package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * Handler for {@code CREATE_SUMMARY}.
 * <p>
 * Produces a deterministic summary sheet from a source sheet. The summary
 * either mirrors the entire source (all columns) or a specified subset.
 * The original sheet is never modified; output goes to a new {@code _summary}
 * sheet in the output workbook.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code sourceSheet} — optional: name of the source sheet; defaults to
 *       the context's active sheet</li>
 *   <li>{@code title} — optional: free-form title (recorded in summary, not
 *       written into the sheet)</li>
 *   <li>{@code columns} — optional: comma-separated list of columns to include;
 *       if omitted, all columns are included</li>
 * </ul>
 */
@Component
@Slf4j
public class CreateSummaryHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "CREATE_SUMMARY";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        Map<String, String> params = step.getParameters() == null
                ? Map.of() : step.getParameters();

        // Resolve source sheet: explicit param > active sheet
        String sourceSheetName = params.get("sourceSheet");
        if (sourceSheetName == null || sourceSheetName.isBlank()) {
            sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        }
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        // ── Read header row ──────────────────────────────────────────────────
        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "CREATE_SUMMARY step " + step.getOrder()
                    + ": source sheet '" + sourceSheetName + "' is empty");
        }
        Row headerRow = sourceSheet.getRow(firstRow);
        if (headerRow == null) {
            throw new WorkflowExecutionException(
                    "CREATE_SUMMARY step " + step.getOrder()
                    + ": source sheet '" + sourceSheetName + "' has no header row");
        }

        int colCount = headerRow.getLastCellNum();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < colCount; c++) {
            Cell hc = headerRow.getCell(c);
            headers.add(hc != null && hc.getCellType() == CellType.STRING
                    ? hc.getStringCellValue().trim() : "");
        }

        // ── Resolve which columns to include ────────────────────────────────
        List<String> requestedColumns = parseColumnsParam(params.get("columns"));

        List<Integer> includeIndices;  // indices in source sheet
        List<String>  includeNames;    // names in the order to write

        if (requestedColumns.isEmpty()) {
            // No columns specified → include all
            includeIndices = new ArrayList<>();
            includeNames = new ArrayList<>();
            for (int c = 0; c < headers.size(); c++) {
                if (!headers.get(c).isEmpty()) {
                    includeIndices.add(c);
                    includeNames.add(headers.get(c));
                }
            }
        } else {
            includeIndices = new ArrayList<>();
            includeNames = new ArrayList<>();
            for (String col : requestedColumns) {
                int idx = headers.indexOf(col);
                if (idx < 0) {
                    throw new WorkflowExecutionException(
                            "CREATE_SUMMARY step " + step.getOrder()
                            + ": column '" + col + "' not found in source sheet '"
                            + sourceSheetName + "'");
                }
                includeIndices.add(idx);
                includeNames.add(col);
            }
        }

        if (includeIndices.isEmpty()) {
            throw new WorkflowExecutionException(
                    "CREATE_SUMMARY step " + step.getOrder()
                    + ": source sheet '" + sourceSheetName + "' has no usable columns");
        }

        // ── Create output sheet ─────────────────────────────────────────────
        String outSheetName = sourceSheetName + "_summary";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_summary_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Write header
        Row outHeader = outSheet.createRow(0);
        for (int c = 0; c < includeNames.size(); c++) {
            outHeader.createCell(c).setCellValue(includeNames.get(c));
        }

        // Copy data rows
        int lastRow = sourceSheet.getLastRowNum();
        int copiedRows = 0;
        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row srcRow = sourceSheet.getRow(r);
            if (srcRow == null) continue;

            Row outRow = outSheet.createRow(outSheet.getLastRowNum() + 1);
            for (int i = 0; i < includeIndices.size(); i++) {
                Cell srcCell = srcRow.getCell(includeIndices.get(i));
                if (srcCell == null) continue;
                Cell outCell = outRow.createCell(i);
                copyCellValue(srcCell, outCell);
            }
            copiedRows++;
        }

        context.setActiveSheet(outSheetName);

        String title = params.get("title");
        String summary = String.format(
                "CREATE_SUMMARY of '%s' → sheet '%s' (%d columns, %d rows)%s",
                sourceSheetName, outSheetName, includeNames.size(), copiedRows,
                (title != null && !title.isBlank() ? " — title: " + title : ""));
        log.debug("CREATE_SUMMARY step {}: {}", step.getOrder(), summary);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .columns(includeNames)
                .rowCount(copiedRows)
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    /**
     * Parses the {@code columns} parameter: comma-separated list, trimmed, blanks removed.
     */
    private List<String> parseColumnsParam(String columnsParam) {
        List<String> result = new ArrayList<>();
        if (columnsParam == null || columnsParam.isBlank()) {
            return result;
        }
        for (String col : columnsParam.split(",")) {
            String trimmed = col.trim();
            if (!trimmed.isEmpty()) {
                if (result.contains(trimmed)) {
                    // duplicates are silently deduplicated to keep output deterministic
                    continue;
                }
                result.add(trimmed);
            }
        }
        return result;
    }

    private void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING  -> target.setCellValue(source.getStringCellValue());
            case NUMERIC -> target.setCellValue(source.getNumericCellValue());
            case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
            case BLANK   -> target.setCellType(CellType.BLANK);
            case FORMULA -> {
                try { target.setCellFormula(source.getCellFormula()); }
                catch (Exception ignored) { target.setCellType(CellType.BLANK); }
            }
            default -> {}
        }
    }
}
