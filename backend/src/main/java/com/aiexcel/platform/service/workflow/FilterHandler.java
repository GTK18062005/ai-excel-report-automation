package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Handler for {@code FILTER}.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code column} — required: the column to test (must exist in source sheet)</li>
 *   <li>{@code operator} — required: one of equals, not_equals, greater_than, less_than,
 *       greater_or_equal, less_or_equal, contains, not_contains, in</li>
 *   <li>{@code value} — the comparison value (string compared as text or as a number
 *       if the cell is numeric)</li>
 *   <li>{@code values} — comma-separated list (only for {@code in})</li>
 * </ul>
 * <p>
 * Rows whose value does not satisfy the predicate are dropped.
 */
@Component
@Slf4j
public class FilterHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "FILTER";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        Map<String, String> params = step.getParameters();
        if (params == null) {
            throw new WorkflowExecutionException(
                    "FILTER step " + step.getOrder() + ": parameters must be provided");
        }

        String column = params.get("column");
        if (column == null || column.isBlank()) {
            column = params.get("columnName");
        }
        if (column == null || column.isBlank()) {
            throw new WorkflowExecutionException(
                    "FILTER step " + step.getOrder() + ": parameter 'column' is required");
        }

        String operator = params.get("operator");
        if (operator == null || operator.isBlank()) {
            throw new WorkflowExecutionException(
                    "FILTER step " + step.getOrder() + ": parameter 'operator' is required");
        }
        operator = operator.trim().toLowerCase();

        if (!SUPPORTED_OPERATORS.contains(operator)) {
            throw new WorkflowExecutionException(
                    "FILTER step " + step.getOrder() + ": unsupported operator '" + operator + "'. Supported: "
                            + String.join(", ", SUPPORTED_OPERATORS));
        }

        String value = params.get("value");
        String valuesRaw = params.get("values");

        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "FILTER step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' is empty");
        }
        Row header = sourceSheet.getRow(firstRow);
        if (header == null) {
            throw new WorkflowExecutionException(
                    "FILTER step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' has no header row");
        }

        int colIndex = -1;
        int colCount = header.getLastCellNum();
        for (int c = 0; c < colCount; c++) {
            Cell hc = header.getCell(c);
            String name = (hc != null && hc.getCellType() == CellType.STRING) ? hc.getStringCellValue().trim() : "";
            if (column.equals(name)) {
                colIndex = c;
                break;
            }
        }
        if (colIndex < 0) {
            throw new WorkflowExecutionException(
                    "FILTER step " + step.getOrder() + ": column '" + column
                            + "' not found in source sheet '" + sourceSheetName + "'");
        }

        // Pre-parse "in" set if needed
        Set<String> inSet = null;
        if ("in".equals(operator)) {
            if (valuesRaw == null || valuesRaw.isBlank()) {
                throw new WorkflowExecutionException(
                        "FILTER step " + step.getOrder() + ": parameter 'values' is required for operator 'in'");
            }
            inSet = new HashSet<>();
            for (String v : valuesRaw.split(",")) {
                String trimmed = v.trim();
                if (!trimmed.isEmpty()) inSet.add(trimmed);
            }
        }

        // Create output sheet
        String outSheetName = sourceSheetName + "_filtered";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_filtered_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Copy header
        copyRow(header, outSheet.createRow(0));

        int lastRow = sourceSheet.getLastRowNum();
        int rowsKept = 0;
        int rowsRemoved = 0;
        int outRowIndex = 1;

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row sourceRow = sourceSheet.getRow(r);
            if (sourceRow == null) continue;
            Cell cell = sourceRow.getCell(colIndex);
            boolean match = evaluate(cell, operator, value, inSet);
            if (!match) {
                rowsRemoved++;
                continue;
            }
            Row outRow = outSheet.createRow(outRowIndex++);
            copyRow(sourceRow, outRow);
            rowsKept++;
        }

        String summary = String.format(
                "Filter %s %s %s → kept %d rows, removed %d → sheet '%s'",
                column, operator, value != null ? "'" + value + "'" : "...",
                rowsKept, rowsRemoved, outSheetName);
        log.debug("FILTER step {}: {}", step.getOrder(), summary);

        context.setActiveSheet(outSheetName);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .rowCount(rowsKept)
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    private boolean evaluate(Cell cell, String operator, String value, Set<String> inSet) {
        if ("in".equals(operator)) {
            String v = cellToComparableString(cell);
            return inSet.contains(v);
        }
        // For numeric comparisons, try numeric compare; otherwise string compare.
        if (cell == null) return false;

        boolean cellIsNumeric = cell.getCellType() == CellType.NUMERIC;
        if (cellIsNumeric && (operator.equals("greater_than") || operator.equals("less_than")
                || operator.equals("greater_or_equal") || operator.equals("less_or_equal"))) {
            double cellNum = cell.getNumericCellValue();
            double cmp;
            try {
                cmp = Double.parseDouble(value);
            } catch (NumberFormatException e) {
                // fall through to string compare
                return false;
            }
            return switch (operator) {
                case "greater_than" -> cellNum > cmp;
                case "less_than" -> cellNum < cmp;
                case "greater_or_equal" -> cellNum >= cmp;
                case "less_or_equal" -> cellNum <= cmp;
                default -> false;
            };
        }

        String cellStr = cellToComparableString(cell);
        String cmpStr = value == null ? "" : value.trim();
        return switch (operator) {
            case "equals" -> cellStr.equalsIgnoreCase(cmpStr);
            case "not_equals" -> !cellStr.equalsIgnoreCase(cmpStr);
            case "contains" -> cellStr.toLowerCase().contains(cmpStr.toLowerCase());
            case "not_contains" -> !cellStr.toLowerCase().contains(cmpStr.toLowerCase());
            case "greater_than" -> safeCompare(cellStr, cmpStr) > 0;
            case "less_than" -> safeCompare(cellStr, cmpStr) < 0;
            case "greater_or_equal" -> safeCompare(cellStr, cmpStr) >= 0;
            case "less_or_equal" -> safeCompare(cellStr, cmpStr) <= 0;
            default -> false;
        };
    }

    private int safeCompare(String a, String b) {
        // Try numeric compare if both look like numbers
        try {
            double da = Double.parseDouble(a);
            double db = Double.parseDouble(b);
            return Double.compare(da, db);
        } catch (NumberFormatException ignored) {
            return a.compareToIgnoreCase(b);
        }
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

    private static final List<String> SUPPORTED_OPERATORS = List.of(
            "equals", "not_equals", "greater_than", "less_than",
            "greater_or_equal", "less_or_equal", "contains", "not_contains", "in");

    private void copyRow(Row source, Row target) {
        int lastCell = source.getLastCellNum();
        for (int c = 0; c < lastCell; c++) {
            Cell src = source.getCell(c);
            if (src == null) continue;
            Cell tgt = target.createCell(c);
            copyCellValue(src, tgt);
        }
    }

    private void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING -> target.setCellValue(source.getStringCellValue());
            case NUMERIC -> target.setCellValue(source.getNumericCellValue());
            case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
            case BLANK -> target.setCellType(CellType.BLANK);
            case FORMULA -> {
                try {
                    target.setCellFormula(source.getCellFormula());
                } catch (Exception ignored) {
                    target.setCellType(CellType.BLANK);
                }
            }
            default -> { /* unsupported, leave blank */ }
        }
    }
}
