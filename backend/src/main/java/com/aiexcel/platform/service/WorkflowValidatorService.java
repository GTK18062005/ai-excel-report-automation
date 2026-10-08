package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.*;
import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.exception.WorkflowValidationException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.*;
import java.util.stream.Collectors;

/**
 * Deterministic workflow validator.
 * <p>
 * Validates a {@link WorkflowProposalResponse} (or its request-body equivalent
 * {@link WorkflowValidationRequest}) before the user approves it for execution.
 * <p>
 * Validation rules:
 * <ul>
 *   <li>Steps list must not be null.</li>
 *   <li>Each step must have a valid operation from the {@link WorkflowOperation} catalog.</li>
 *   <li>Step {@code order} values must be sequential positive integers starting at 1.</li>
 *   <li>Referenced sheet names must exist in the target workbook.</li>
 *   <li>Referenced column names must exist in their respective sheets.</li>
 *   <li>Required operation fields (operation name, and inputSheets for most ops) must be present.</li>
 *   <li>Unsafe/unknown operations are rejected.</li>
 *   <li>Parameters map must not be null (may be empty).</li>
 * </ul>
 * <p>
 * This service performs NO AI calls, NO workflow execution, NO file modification.
 */
@Service
@RequiredArgsConstructor
@Slf4j
public class WorkflowValidatorService {

    private final UploadedFileRepository fileRepository;

    /**
     * Required parameter keys per operation. Presence-only check; value-shape
     * validation is the responsibility of the individual handler at execution time.
     * <p>
     * Operations not listed here (READ_SHEET, SELECT_COLUMNS, REMOVE_EMPTY_ROWS,
     * REMOVE_DUPLICATES, CREATE_SUMMARY, VALIDATE, GENERATE_PDF, EXPORT_EXCEL) take
     * no parameters.
     */
    private static final Map<String, List<String>> REQUIRED_PARAMETERS = Map.of(
            "GROUP_BY",     List.of("groupColumns"),
            "FILTER",       List.of("column", "operator", "value"),
            "SORT",         List.of("column", "direction"),
            "AGGREGATE",    List.of("aggregateColumn", "aggregateFunction"),
            "PIVOT",        List.of("rowColumn", "pivotColumn", "valueColumn", "aggregationFunction"),
            "CALCULATE",    List.of("targetColumn", "expression"),
            "JOIN_SHEETS",  List.of("leftSheet", "rightSheet", "joinColumn", "joinType"),
            "CREATE_CHART", List.of("sourceSheet", "chartType", "labelColumn", "dataColumn")
    );

    /**
     * Validates a workflow proposal against the target workbook.
     *
     * @param fileId   ID of the uploaded file (ownership is verified)
     * @param user     authenticated user (must own the file)
     * @param request  workflow proposal body to validate
     * @return validation result — never null
     */
    @Transactional(readOnly = true)
    public WorkflowValidationResponse validateWorkflow(Long fileId, User user, WorkflowValidationRequest request) {
        log.info("Validating workflow for fileId={}, userId={}", fileId, user.getId());

        // 1. Verify file ownership
        UploadedFile file = fileRepository
                .findByIdAndUserAndStatusNot(fileId, user, FileStatus.DELETED)
                .orElseThrow(() -> new FileNotFoundException("File not found with id: " + fileId));

        List<String> errors = new ArrayList<>();

        // 2. Steps null check
        List<WorkflowStep> steps = request.getSteps();
        if (steps == null) {
            errors.add("Steps list must not be null");
            return WorkflowValidationResponse.builder()
                    .valid(false)
                    .errors(errors)
                    .build();
        }

        // 3. Step-level validations (order, operation, required fields, unsafe operations)
        validateStepLevelErrors(steps, errors);

        // 4. Structural validations (sequential order)
        validateStepOrder(steps, errors);

        // 5. Resolve sheet and column names from the workbook
        Map<String, List<String>> sheetColumns = resolveSheetColumns(file);
        if (sheetColumns == null) {
            // File could not be read — service-layer exception already logged inside resolveSheetColumns
            throw new WorkflowValidationException("Unable to read workbook to resolve sheet and column names");
        }

        // 6. Cross-reference validations (sheet existence, column existence)
        validateReferences(steps, sheetColumns, errors);

        boolean valid = errors.isEmpty();
        log.info("Workflow validation complete for fileId={}, valid={}, errorCount={}",
                fileId, valid, errors.size());

        return WorkflowValidationResponse.builder()
                .valid(valid)
                .errors(errors)
                .build();
    }

    /**
     * Validates each step's fields independently (no workbook context required).
     */
    private void validateStepLevelErrors(List<WorkflowStep> steps, List<String> errors) {
        for (int i = 0; i < steps.size(); i++) {
            WorkflowStep step = steps.get(i);
            int oneBased = i + 1; // user-facing step number

            // Step must not be null
            if (step == null) {
                errors.add("Step " + oneBased + ": step must not be null");
                continue;
            }

            // Operation name must be present
            if (step.getOperation() == null || step.getOperation().isBlank()) {
                errors.add("Step " + oneBased + ": operation must not be blank");
            } else {
                // Operation must be in the allowed catalog
                if (!WorkflowOperation.isValid(step.getOperation())) {
                    errors.add("Step " + oneBased + ": unknown or unsafe operation '" + step.getOperation() + "' — only operations from the allowed catalog are permitted");
                }

                // Validate required fields per operation type
                validateRequiredFields(step, oneBased, errors);

                // Validate required parameter keys per operation type
                validateRequiredParameters(step, oneBased, errors);
            }

            // inputSheets must not be null (empty is acceptable)
            if (step.getInputSheets() == null) {
                errors.add("Step " + oneBased + ": inputSheets must not be null");
            }

            // inputColumns must not be null (empty is acceptable)
            if (step.getInputColumns() == null) {
                errors.add("Step " + oneBased + ": inputColumns must not be null");
            }

            // Parameters map must not be null (empty is acceptable)
            if (step.getParameters() == null) {
                errors.add("Step " + oneBased + ": parameters map must not be null");
            }
        }
    }

    /**
     * Validates required fields per operation type.
     * Currently enforces: most operations require at least one input sheet.
     */
    private void validateRequiredFields(WorkflowStep step, int oneBased, List<String> errors) {
        List<String> inputSheets = step.getInputSheets();

        if (inputSheets == null || inputSheets.isEmpty()) {
            String op = step.getOperation();
            // All non-trivial operations require at least one input sheet
            if (!"GENERATE_PDF".equals(op) && !"EXPORT_EXCEL".equals(op) && !"CREATE_CHART".equals(op)) {
                errors.add("Step " + oneBased + ": operation '" + op + "' requires at least one input sheet");
            }
        }
    }

    /**
     * Validates that every required parameter key for this operation is present
     * in {@code step.parameters}. A key is considered present only if its value
     * is non-null and not blank. Value-shape validation is the handler's job.
     */
    private void validateRequiredParameters(WorkflowStep step, int oneBased, List<String> errors) {
        String op = step.getOperation();
        List<String> required = REQUIRED_PARAMETERS.get(op);
        Map<String, String> params = step.getParameters();

        if (required != null && !required.isEmpty()) {
            for (String key : required) {
                String value = params == null ? null : params.get(key);
                if (value == null || value.isBlank()) {
                    errors.add("Step " + oneBased + ": " + op + " requires parameter '" + key + "'");
                }
            }
        }

        // Validate optional parameters that must not be blank if provided
        if ("AGGREGATE".equals(op) && params != null && params.containsKey("groupColumns")) {
            String value = params.get("groupColumns");
            if (value == null || value.isBlank()) {
                errors.add("Step " + oneBased + ": " + op + " parameter 'groupColumns' must not be blank if provided");
            }
        }
    }

    /**
     * Validates that step order values are sequential positive integers starting at 1.
     */
    private void validateStepOrder(List<WorkflowStep> steps, List<String> errors) {
        for (int i = 0; i < steps.size(); i++) {
            WorkflowStep step = steps.get(i);
            if (step == null) continue; // already reported above

            int expected = i + 1;
            if (step.getOrder() != expected) {
                errors.add("Step " + (i + 1) + ": step order must be " + expected + " but was " + step.getOrder());
            }
            if (step.getOrder() <= 0) {
                errors.add("Step " + (i + 1) + ": step order must be a positive integer, got " + step.getOrder());
            }
        }
    }

    /**
     * Resolves sheet names and their column names from the workbook.
     * Returns a map of sheetName -> list of column names.
     * Returns null if the workbook cannot be read.
     */
    private Map<String, List<String>> resolveSheetColumns(UploadedFile file) {
        String fileType = file.getFileType();
        if (fileType == null ||
            (!"xlsx".equalsIgnoreCase(fileType) && !"xls".equalsIgnoreCase(fileType))) {
            log.debug("File type {} does not support sheet/column resolution", fileType);
            return null;
        }

        java.nio.file.Path path = java.nio.file.Paths.get(file.getStoragePath());
        if (!java.nio.file.Files.exists(path)) {
            log.warn("Physical file not found at path: {}", file.getStoragePath());
            return null;
        }

        try (org.apache.poi.ss.usermodel.Workbook workbook =
                org.apache.poi.ss.usermodel.WorkbookFactory.create(path.toFile(), null, true)) {

            Map<String, List<String>> result = new HashMap<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                if (sheetName == null) sheetName = "Sheet" + (i + 1);

                List<String> columns = resolveSheetColumnsFromSheet(sheet);
                result.put(sheetName, columns);
            }
            return result;

        } catch (Exception e) {
            log.warn("Failed to read workbook for sheet/column resolution: {}", e.getMessage());
            return null;
        }
    }

    private List<String> resolveSheetColumnsFromSheet(org.apache.poi.ss.usermodel.Sheet sheet) {
        int firstRow = sheet.getFirstRowNum();
        if (firstRow < 0) return List.of();

        org.apache.poi.ss.usermodel.Row headerRow = sheet.getRow(firstRow);
        if (headerRow == null) return List.of();

        int lastCell = headerRow.getLastCellNum();
        List<String> columns = new ArrayList<>(lastCell);
        for (int col = 0; col < lastCell; col++) {
            org.apache.poi.ss.usermodel.Cell cell = headerRow.getCell(col);
            String value = getCellValueAsString(cell);
            columns.add(value != null ? value.trim() : "");
        }
        return columns;
    }

    private String getCellValueAsString(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue() != null ? cell.getStringCellValue() : "";
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    /**
     * Validates sheet and column references for each step.
     */
    private void validateReferences(List<WorkflowStep> steps,
                                   Map<String, List<String>> sheetColumns,
                                   List<String> errors) {
        // Build a flat map of columnName -> sheetName for fast duplicate-column lookup
        Map<String, String> columnToSheet = buildColumnToSheetMap(sheetColumns);

        for (int i = 0; i < steps.size(); i++) {
            WorkflowStep step = steps.get(i);
            if (step == null) continue;
            if (!WorkflowOperation.isValid(step.getOperation())) continue; // already reported above
            if (step.getInputSheets() == null) continue; // already reported above

            int oneBased = i + 1;

            // Validate sheet existence
            for (String sheetName : step.getInputSheets()) {
                if (sheetName == null || sheetName.isBlank()) {
                    errors.add("Step " + oneBased + ": input sheet name must not be blank");
                    continue;
                }
                if (!sheetColumns.containsKey(sheetName)) {
                    errors.add("Step " + oneBased + ": referenced sheet '" + sheetName + "' does not exist in workbook");
                }
            }

            // Validate column existence within referenced sheets
            if (step.getInputColumns() != null) {
                for (String columnName : step.getInputColumns()) {
                    if (columnName == null || columnName.isBlank()) {
                        errors.add("Step " + oneBased + ": input column name must not be blank");
                        continue;
                    }

                    boolean found = false;
                    for (String sheetName : step.getInputSheets()) {
                        if (sheetName == null || sheetName.isBlank()) continue;
                        List<String> sheetColList = sheetColumns.get(sheetName);
                        if (sheetColList != null && sheetColList.contains(columnName)) {
                            found = true;
                            break;
                        }
                    }
                    if (!found) {
                        errors.add("Step " + oneBased + ": referenced column '" + columnName
                                + "' not found in any of the input sheets ["
                                + String.join(", ", step.getInputSheets()) + "]");
                    }
                }
            }
        }
    }

    /**
     * Builds a map of columnName -> sheetName for detecting columns that exist
     * across multiple sheets.
     */
    private Map<String, String> buildColumnToSheetMap(Map<String, List<String>> sheetColumns) {
        Map<String, String> map = new HashMap<>();
        for (Map.Entry<String, List<String>> entry : sheetColumns.entrySet()) {
            for (String col : entry.getValue()) {
                map.put(col, entry.getKey());
            }
        }
        return map;
    }
}
