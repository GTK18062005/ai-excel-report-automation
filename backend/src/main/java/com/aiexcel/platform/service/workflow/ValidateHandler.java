package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Component;

import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * Handler for {@code VALIDATE}.
 * <p>
 * A non-transformative workflow step that checks the source sheet for
 * data-quality conditions and records the result in the step summary.
 * This handler NEVER modifies the input workbook or creates a new sheet.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code targetColumn} — required: the column to validate (must exist in source sheet)</li>
 *   <li>{@code checkType} — required: one of not_null, unique, positive</li>
 * </ul>
 * <p>
 * Validation failure causes the step to throw a {@link WorkflowExecutionException},
 * which stops the workflow execution chain. This is the same abort-on-failure
 * convention used by all other handlers.
 * <p>
 * The source sheet is resolved via {@link ExecutionUtils#resolveSourceSheet}
 * (inputSheets first, then activeSheet), matching the convention of other handlers.
 */
@Component
@Slf4j
public class ValidateHandler implements WorkflowOperationHandler {

    private static final String NOT_NULL = "not_null";
    private static final String UNIQUE = "unique";
    private static final String POSITIVE = "positive";

    @Override
    public String getSupportedOperation() {
        return "VALIDATE";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        Map<String, String> params = step.getParameters();
        if (params == null) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + step.getOrder() + ": parameters must be provided");
        }

        String targetColumn = resolveColumn(params);
        if (targetColumn == null || targetColumn.isBlank()) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + step.getOrder() + ": parameter 'targetColumn' is required");
        }

        String checkTypeRaw = params.get("checkType");
        if (checkTypeRaw == null || checkTypeRaw.isBlank()) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + step.getOrder() + ": parameter 'checkType' is required");
        }
        String checkType = checkTypeRaw.trim().toLowerCase();
        if (!SUPPORTED_CHECK_TYPES.contains(checkType)) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + step.getOrder() + ": unsupported checkType '" + checkType
                            + "'. Supported: " + String.join(", ", SUPPORTED_CHECK_TYPES));
        }

        int colIndex = resolveColumnIndex(sourceSheet, targetColumn, step.getOrder());

        // Update active sheet so subsequent steps can chain from here
        context.setActiveSheet(sourceSheetName);

        // Delegate to the specific check
        return switch (checkType) {
            case NOT_NULL -> checkNotNull(step, sourceSheet, colIndex, targetColumn, sourceSheetName);
            case UNIQUE -> checkUnique(step, sourceSheet, colIndex, targetColumn, sourceSheetName);
            case POSITIVE -> checkPositive(step, sourceSheet, colIndex, targetColumn, sourceSheetName);
            default -> throw new WorkflowExecutionException(
                    "VALIDATE step " + step.getOrder() + ": unsupported checkType '" + checkType + "'");
        };
    }

    private StepExecutionResult checkNotNull(WorkflowStep step, Sheet sheet, int colIndex,
                                            String column, String sheetName) {
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

        int nullCount = 0;
        int dataRowCount = 0;

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(colIndex);
            dataRowCount++;
            if (isBlank(cell)) {
                nullCount++;
            }
        }

        String summary = String.format(
                "Column '%s' — %d null/blank values found in %d data rows → %s",
                column, nullCount, dataRowCount,
                nullCount == 0 ? "PASS" : "FAIL");
        log.debug("VALIDATE (not_null) step {}: {}", step.getOrder(), summary);

        if (nullCount > 0) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + step.getOrder() + ": column '" + column
                            + "' contains " + nullCount + " null/blank value(s) → validation FAILED");
        }

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(sheetName)
                .rowCount(dataRowCount)
                .summary(summary)
                .nextActiveSheet(sheetName)
                .build();
    }

    private StepExecutionResult checkUnique(WorkflowStep step, Sheet sheet, int colIndex,
                                           String column, String sheetName) {
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

        Set<String> seen = new HashSet<>();
        Set<String> duplicates = new HashSet<>();
        int dataRowCount = 0;

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(colIndex);
            String value = cellToComparableString(cell);
            dataRowCount++;
            if (!value.isEmpty()) {
                if (!seen.add(value)) {
                    duplicates.add(value);
                }
            }
        }

        String summary = String.format(
                "Column '%s' — %d unique values, %d duplicates found → %s",
                column, seen.size(), duplicates.size(),
                duplicates.isEmpty() ? "PASS" : "FAIL");
        log.debug("VALIDATE (unique) step {}: {}", step.getOrder(), summary);

        if (!duplicates.isEmpty()) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + step.getOrder() + ": column '" + column
                            + "' contains " + duplicates.size() + " duplicate value(s) → validation FAILED");
        }

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(sheetName)
                .rowCount(dataRowCount)
                .summary(summary)
                .nextActiveSheet(sheetName)
                .build();
    }

    private StepExecutionResult checkPositive(WorkflowStep step, Sheet sheet, int colIndex,
                                              String column, String sheetName) {
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

        int nonPositiveCount = 0;
        int dataRowCount = 0;

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(colIndex);
            dataRowCount++;
            if (cell == null || cell.getCellType() != CellType.NUMERIC) {
                // Treat null or non-numeric as non-positive (zero or invalid)
                nonPositiveCount++;
            } else {
                double value = cell.getNumericCellValue();
                if (value <= 0) {
                    nonPositiveCount++;
                }
            }
        }

        String summary = String.format(
                "Column '%s' — %d positive values, %d non-positive found → %s",
                column, dataRowCount - nonPositiveCount, nonPositiveCount,
                nonPositiveCount == 0 ? "PASS" : "FAIL");
        log.debug("VALIDATE (positive) step {}: {}", step.getOrder(), summary);

        if (nonPositiveCount > 0) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + step.getOrder() + ": column '" + column
                            + "' contains " + nonPositiveCount + " non-positive value(s) → validation FAILED");
        }

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(sheetName)
                .rowCount(dataRowCount)
                .summary(summary)
                .nextActiveSheet(sheetName)
                .build();
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    private String resolveColumn(Map<String, String> params) {
        String col = params.get("targetColumn");
        if (col != null && !col.isBlank()) return col.trim();
        col = params.get("column");
        if (col != null && !col.isBlank()) return col.trim();
        return null;
    }

    private int resolveColumnIndex(Sheet sheet, String targetColumn, int stepOrder) {
        int firstRow = sheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + stepOrder + ": source sheet is empty");
        }
        Row header = sheet.getRow(firstRow);
        if (header == null) {
            throw new WorkflowExecutionException(
                    "VALIDATE step " + stepOrder + ": source sheet has no header row");
        }

        int colCount = header.getLastCellNum();
        for (int c = 0; c < colCount; c++) {
            Cell hc = header.getCell(c);
            String name = cellToComparableString(hc).trim();
            if (targetColumn.equals(name)) {
                return c;
            }
        }
        throw new WorkflowExecutionException(
                "VALIDATE step " + stepOrder + ": column '" + targetColumn
                        + "' not found in source sheet");
    }

    private boolean isBlank(Cell cell) {
        if (cell == null) return true;
        return switch (cell.getCellType()) {
            case BLANK, _NONE -> true;
            case STRING -> cell.getStringCellValue() == null
                    || cell.getStringCellValue().isBlank();
            default -> false;
        };
    }

    private String cellToComparableString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> {
                String v = cell.getStringCellValue();
                yield v != null ? v.trim() : "";
            }
            case NUMERIC -> {
                double d = cell.getNumericCellValue();
                if (d == Math.floor(d) && !Double.isInfinite(d)) {
                    yield String.valueOf((long) d);
                }
                yield String.valueOf(d);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    private static final java.util.List<String> SUPPORTED_CHECK_TYPES =
            java.util.List.of(NOT_NULL, UNIQUE, POSITIVE);
}
