package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowExecutionResponse;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.entity.WorkflowApproval;
import com.aiexcel.platform.entity.WorkflowStatus;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.exception.WorkflowApprovalException;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import com.aiexcel.platform.repository.WorkflowApprovalRepository;
import com.aiexcel.platform.service.workflow.ExecutionContext;
import com.aiexcel.platform.service.workflow.ExecutionUtils;
import com.aiexcel.platform.service.workflow.WorkflowOperationHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.*;
import java.nio.file.*;
import java.util.*;

/**
 * Executes approved workflows against an uploaded workbook.
 * <p>
 * Execution is ONLY allowed when the workflow's {@link WorkflowStatus} is {@code APPROVED}.
 * PROPOSED and REJECTED workflows are rejected with {@link WorkflowApprovalException}.
 * <p>
 * This service performs NO AI calls and NO AI-generated code execution.
 * All operations are deterministic Java transformations via Apache POI.
 * <p>
 * The original uploaded workbook is never modified; results are written to a new
 * output file in the same storage directory.
 */
@Service
@Slf4j
public class WorkflowExecutionService {

    private final WorkflowApprovalRepository approvalRepository;
    private final UploadedFileRepository fileRepository;
    private final FileStorageService fileStorageService;
    private final Map<String, WorkflowOperationHandler> handlers;

    public WorkflowExecutionService(
            WorkflowApprovalRepository approvalRepository,
            UploadedFileRepository fileRepository,
            FileStorageService fileStorageService,
            List<WorkflowOperationHandler> handlerList
    ) {
        this.approvalRepository = approvalRepository;
        this.fileRepository = fileRepository;
        this.fileStorageService = fileStorageService;
        this.handlers = new HashMap<>();
        for (WorkflowOperationHandler h : handlerList) {
            this.handlers.put(h.getSupportedOperation(), h);
        }
        log.info("WorkflowExecutionService initialized with {} handlers: {}",
                handlers.size(), handlers.keySet());
    }

    /**
     * Executes the approved workflow for the given file.
     * <p>
     * Guard rules:
     * <ul>
     *   <li>User must be authenticated (caller responsibility).</li>
     *   <li>File must exist and be owned by the user.</li>
     *   <li>Workflow status must be {@code APPROVED}; PROPOSED and REJECTED are rejected.</li>
     *   <li>All steps must be supported operations.</li>
     * </ul>
     *
     * @param fileId ID of the uploaded file
     * @param user   authenticated user (must own the file)
     * @return execution result, never null
     */
    @Transactional(readOnly = true)
    public WorkflowExecutionResponse executeWorkflow(Long fileId, User user) {
        log.info("Starting workflow execution for fileId={}, userId={}", fileId, user.getId());

        // 1. Verify file ownership
        UploadedFile file = fileRepository
                .findByIdAndUserAndStatusNot(fileId, user, FileStatus.DELETED)
                .orElseThrow(() -> new FileNotFoundException("File not found with id: " + fileId));

        // 2. Verify physical file exists
        Path sourcePath = Paths.get(file.getStoragePath());
        if (!Files.exists(sourcePath)) {
            throw new WorkflowExecutionException(
                    "Physical file not found for id: " + fileId + " — cannot execute workflow");
        }

        // 3. Fetch and validate approval record
        WorkflowApproval approval = approvalRepository.findByFileAndUser(file, user)
                .orElseThrow(() -> new WorkflowApprovalException(
                        "No workflow proposal found for file id: " + fileId));

        if (approval.getStatus() == WorkflowStatus.PROPOSED) {
            throw new WorkflowApprovalException(
                    "Workflow is not approved yet (status: PROPOSED). Please approve the workflow before executing it.");
        }

        if (approval.getStatus() == WorkflowStatus.REJECTED) {
            throw new WorkflowApprovalException(
                    "Workflow has been rejected and cannot be executed.");
        }

        // 4. Load steps
        List<WorkflowStep> steps = approval.getStepsFromJson();
        if (steps == null || steps.isEmpty()) {
            throw new WorkflowExecutionException("Workflow has no steps to execute");
        }

        // 5. Validate all steps are supported
        validateSupportedOperations(steps);

        // 6. Execute
        return doExecute(sourcePath, file, steps);
    }

    private void validateSupportedOperations(List<WorkflowStep> steps) {
        for (WorkflowStep step : steps) {
            if (step == null) continue;
            String op = step.getOperation();
            if (op == null || op.isBlank()) {
                throw new WorkflowExecutionException("Step " + step.getOrder() + " has a blank operation name");
            }
            if (!handlers.containsKey(op)) {
                throw new WorkflowExecutionException(
                        "Step " + step.getOrder() + ": operation '" + op
                                + "' is not yet implemented for execution. "
                                + "Currently supported: " + String.join(", ", handlers.keySet()));
            }
        }
    }

    private WorkflowExecutionResponse doExecute(Path sourcePath, UploadedFile file, List<WorkflowStep> steps) {
        String fileType = file.getFileType();
        if (!"xlsx".equalsIgnoreCase(fileType) && !"xls".equalsIgnoreCase(fileType)) {
            throw new WorkflowExecutionException(
                    "Workflow execution is only supported for .xlsx and .xls files, got: ." + fileType);
        }

        // Derive output path using the same shared naming rule that the
        // download endpoint will use. Keeping this in one place ensures the
        // file the executor writes and the file the endpoint resolves are
        // always the same name.
        Path outputPath = ExecutionUtils.deriveOutputFileName(sourcePath);

        List<WorkflowExecutionResponse.StepResult> stepResults = new ArrayList<>();
        List<String> errors = new ArrayList<>();

        // Open input workbook in read-only mode. The output workbook is a fresh
        // in-memory XSSFWorkbook so the source file is NEVER modified.
        try (
                Workbook inputWorkbook = WorkbookFactory.create(sourcePath.toFile(), null, true);
                Workbook outputWorkbook = newXSSFWorkbook(fileType)
        ) {
            ExecutionContext context = ExecutionContext.builder()
                    .inputWorkbook(inputWorkbook)
                    .outputWorkbook(outputWorkbook)
                    .sourcePath(sourcePath.toString())
                    .outputPath(outputPath.toString())
                    .steps(steps)
                    .activeSheet(null)
                    .build();

            // Preserve all original source sheets in the output workbook so the output
            // is a superset: original sheets + handler-derived sheets. This ensures
            // ExecutionUtils.findSheet() can find source sheets for downstream handlers.
            copySourceSheetsToOutput(inputWorkbook, outputWorkbook);

            for (WorkflowStep step : steps) {
                if (step == null) {
                    errors.add("Encountered a null step — skipping");
                    continue;
                }
                String op = step.getOperation();
                WorkflowOperationHandler handler = handlers.get(op);
                if (handler == null) {
                    // Should not reach here due to validateSupportedOperations
                    errors.add("Step " + step.getOrder() + ": no handler for operation '" + op + "'");
                    continue;
                }

                try {
                    StepExecutionResult result = handler.execute(step, context);
                    stepResults.add(WorkflowExecutionResponse.StepResult.builder()
                            .order(result.getOrder())
                            .operation(result.getOperation())
                            .description(step.getDescription())
                            .resultSummary(result.getSummary())
                            .build());
                    log.debug("Step {} ({}): {}", step.getOrder(), op, result.getSummary());
                } catch (WorkflowExecutionException e) {
                    String msg = "Step " + step.getOrder() + " (" + op + "): " + e.getMessage();
                    log.warn("Execution failed at step {}: {}", step.getOrder(), e.getMessage());
                    errors.add(msg);
                    // Stop on first failure
                    break;
                } catch (Exception e) {
                    String msg = "Step " + step.getOrder() + " (" + op + "): unexpected error — " + e.getMessage();
                    log.error("Unexpected error in step {}", step.getOrder(), e);
                    errors.add(msg);
                    break;
                }
            }

            // Write output workbook
            if (errors.isEmpty()) {
                try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
                    outputWorkbook.write(fos);
                }
                log.info("Workflow execution complete. Output written to: {}", outputPath);
            }

        } catch (IOException e) {
            throw new WorkflowExecutionException("Failed to open or write workbook: " + e.getMessage(), e);
        }

        boolean success = errors.isEmpty();

        if (!success) {
            // Clean up partial output file if it exists
            try {
                Files.deleteIfExists(outputPath);
            } catch (IOException ignored) {}
        }

        return WorkflowExecutionResponse.builder()
                .success(success)
                .message(success
                        ? "Workflow executed successfully"
                        : "Workflow execution failed: " + errors.get(0))
                .outputFileName(success ? outputPath.getFileName().toString() : null)
                .steps(success ? stepResults : null)
                .errors(errors.isEmpty() ? null : errors)
                .build();
    }

    /**
     * Copies every sheet from the input workbook into the output workbook.
     * <p>
     * The original input file is never modified.
     */
    private void copySourceSheetsToOutput(Workbook input, Workbook output) {
        int count = input.getNumberOfSheets();
        for (int i = 0; i < count; i++) {
            Sheet src = input.getSheetAt(i);
            String desiredName = src.getSheetName();
            String targetName = resolveUniqueSheetName(output, desiredName);

            if (output instanceof XSSFWorkbook && src instanceof XSSFSheet) {
                copyXssfSheet((XSSFWorkbook) output, (XSSFSheet) src, targetName);
            } else {
                copySheetValuesManually(src, output, targetName);
            }
        }
    }

    /**
     * Copies a single XSSF sheet into the target XSSF workbook: cell values,
     * formulas, merged regions, column widths, and row heights are preserved.
     * Cell styles and existing drawings/charts are not copied (handlers read
     * source data from the input workbook; CREATE_CHART generates new charts).
     */
    private void copyXssfSheet(XSSFWorkbook tgtWb, XSSFSheet src, String tgtName) {
        Sheet tgt = tgtWb.createSheet(tgtName);

        // Column widths — determine the last used column by scanning rows,
        // since org.apache.poi.ss.usermodel.Sheet has no getLastColumnNum() in 5.2.5.
        int srcLastCol = computeLastColumn(src);
        for (int c = 0; c <= srcLastCol; c++) {
            int w = src.getColumnWidth(c);
            if (w > 0) tgt.setColumnWidth(c, w);
        }

        // Rows and cells
        int firstRow = src.getFirstRowNum();
        int lastRow  = src.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row srcRow = src.getRow(r);
            if (srcRow == null) continue;
            Row tgtRow = tgt.createRow(r);
            tgtRow.setHeight(srcRow.getHeight());
            int firstCell = srcRow.getFirstCellNum();
            int lastCell  = srcRow.getLastCellNum();
            for (int c = firstCell; c < lastCell; c++) {
                Cell srcCell = srcRow.getCell(c);
                if (srcCell == null) continue;
                Cell tgtCell = tgtRow.createCell(c);
                copyCellValue(srcCell, tgtCell);
            }
        }

        // Merged regions
        for (int m = 0; m < src.getNumMergedRegions(); m++) {
            tgt.addMergedRegion(src.getMergedRegion(m));
        }
    }

    /**
     * Returns a sheet name that does not already exist in the output workbook,
     * appending " (copy)" with a numeric suffix if needed.
     */
    private String resolveUniqueSheetName(Workbook output, String desired) {
        if (output.getSheet(desired) == null) return desired;
        int suffix = 2;
        String candidate;
        do {
            candidate = desired + " (" + suffix++ + ")";
        } while (output.getSheet(candidate) != null);
        return candidate;
    }

    /**
     * Cell-by-cell copy: values, formulas, date values, merged regions,
     * column widths, and row heights are copied. Cell styles and drawings
     * are not copied.
     */
    private void copySheetValuesManually(Sheet src, Workbook output, String tgtName) {
        Sheet tgt = output.createSheet(tgtName);

        // Column widths — determine the last used column by scanning rows,
        // since org.apache.poi.ss.usermodel.Sheet has no getLastColumnNum() in 5.2.5.
        int srcLastCol = computeLastColumn(src);
        for (int c = 0; c <= srcLastCol; c++) {
            int w = src.getColumnWidth(c);
            if (w > 0) tgt.setColumnWidth(c, w);
        }

        // Rows and cells
        int firstRow = src.getFirstRowNum();
        int lastRow  = src.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row srcRow = src.getRow(r);
            if (srcRow == null) continue;
            Row tgtRow = tgt.createRow(r);
            tgtRow.setHeight(srcRow.getHeight());
            int firstCell = srcRow.getFirstCellNum();
            int lastCell  = srcRow.getLastCellNum();
            for (int c = firstCell; c < lastCell; c++) {
                Cell srcCell = srcRow.getCell(c);
                if (srcCell == null) continue;
                Cell tgtCell = tgtRow.createCell(c);
                copyCellValue(srcCell, tgtCell);
            }
        }

        // Merged regions
        for (int m = 0; m < src.getNumMergedRegions(); m++) {
            tgt.addMergedRegion(src.getMergedRegion(m));
        }
    }

    /**
     * Copies a single cell's value from source to target.
     * The target cell must already be created by the caller.
     */
    private void copyCellValue(Cell src, Cell tgt) {
        switch (src.getCellType()) {
            case STRING  -> tgt.setCellValue(src.getStringCellValue());
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(src)) {
                    tgt.setCellValue(src.getDateCellValue());
                } else {
                    tgt.setCellValue(src.getNumericCellValue());
                }
            }
            case BOOLEAN -> tgt.setCellValue(src.getBooleanCellValue());
            case FORMULA  -> tgt.setCellFormula(src.getCellFormula());
            case ERROR    -> tgt.setCellErrorValue(src.getErrorCellValue());
            case BLANK, _NONE -> {}
        }
    }

    /**
     * Determines the highest column index used in the given sheet.
     * Scans all rows and returns the maximum last-cell index found.
     * Returns -1 when the sheet is empty.
     */
    private int computeLastColumn(Sheet src) {
        int max = -1;
        int firstRow = src.getFirstRowNum();
        int lastRow  = src.getLastRowNum();
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = src.getRow(r);
            if (row != null) {
                int rowLast = row.getLastCellNum() - 1;
                if (rowLast > max) max = rowLast;
            }
        }
        return max;
    }

    /**
     * Creates a fresh in-memory workbook of the given type so the source file
     * is never opened in write mode.
     */
    private Workbook newXSSFWorkbook(String fileType) {
        if ("xls".equalsIgnoreCase(fileType)) {
            return new org.apache.poi.hssf.usermodel.HSSFWorkbook();
        }
        // default to .xlsx
        return new org.apache.poi.xssf.usermodel.XSSFWorkbook();
    }
}
