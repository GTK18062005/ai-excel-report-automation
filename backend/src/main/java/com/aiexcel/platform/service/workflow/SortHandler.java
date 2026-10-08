package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * Handler for {@code SORT}.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code column} — required: the column to sort by (must exist in source sheet)</li>
 *   <li>{@code direction} — optional: {@code asc} (default) or {@code desc}</li>
 *   <li>{@code type} — optional: {@code text} (default, case-insensitive) or {@code numeric}</li>
 * </ul>
 */
@Component
@Slf4j
public class SortHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "SORT";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        Map<String, String> params = step.getParameters();
        if (params == null) {
            params = Map.of();
        }

        String column = params.get("column");
        if (column == null || column.isBlank()) {
            column = params.get("columnName");
        }
        if (column == null || column.isBlank()) {
            throw new WorkflowExecutionException(
                    "SORT step " + step.getOrder() + ": parameter 'column' is required");
        }

        String directionStr = params.getOrDefault("direction", "asc").trim().toLowerCase();
        if (!directionStr.equals("asc") && !directionStr.equals("desc")) {
            throw new WorkflowExecutionException(
                    "SORT step " + step.getOrder() + ": parameter 'direction' must be 'asc' or 'desc', got '" + directionStr + "'");
        }
        boolean ascending = "asc".equals(directionStr);

        String typeStr = params.getOrDefault("type", "text").trim().toLowerCase();
        if (!typeStr.equals("text") && !typeStr.equals("numeric")) {
            throw new WorkflowExecutionException(
                    "SORT step " + step.getOrder() + ": parameter 'type' must be 'text' or 'numeric', got '" + typeStr + "'");
        }
        boolean numeric = "numeric".equals(typeStr);

        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "SORT step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' is empty");
        }
        Row header = sourceSheet.getRow(firstRow);
        if (header == null) {
            throw new WorkflowExecutionException(
                    "SORT step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' has no header row");
        }

        int colCount = header.getLastCellNum();
        int sortColIndex = -1;
        for (int c = 0; c < colCount; c++) {
            Cell hc = header.getCell(c);
            String name = (hc != null && hc.getCellType() == CellType.STRING) ? hc.getStringCellValue().trim() : "";
            if (column.equals(name)) {
                sortColIndex = c;
                break;
            }
        }
        if (sortColIndex < 0) {
            throw new WorkflowExecutionException(
                    "SORT step " + step.getOrder() + ": column '" + column
                            + "' not found in source sheet '" + sourceSheetName + "'");
        }

        // Read data rows into memory
        int lastRow = sourceSheet.getLastRowNum();
        List<RowData> dataRows = new ArrayList<>();
        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row row = sourceSheet.getRow(r);
            if (row == null) continue;
            // Skip fully-empty rows
            if (isRowEmpty(row)) continue;
            dataRows.add(new RowData(row));
        }

        // Sort — capture in final locals for the lambda
        final int sortCol = sortColIndex;
        final boolean asc = ascending;
        final boolean isNumeric = numeric;
        Comparator<RowData> comparator;
        if (isNumeric) {
            double dir = asc ? 1.0 : -1.0;
            comparator = (a, b) -> {
                double av = a.getNumeric(sortCol);
                double bv = b.getNumeric(sortCol);
                return (int) Math.signum((av - bv) * dir);
            };
        } else {
            comparator = (a, b) -> {
                String av = a.getString(sortCol).toLowerCase();
                String bv = b.getString(sortCol).toLowerCase();
                return asc ? av.compareTo(bv) : bv.compareTo(av);
            };
        }
        dataRows.sort(comparator);

        // Create output sheet
        String outSheetName = sourceSheetName + "_sorted";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_sorted_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Copy header
        copyRow(header, outSheet.createRow(0));

        // Write sorted rows
        int outRowIndex = 1;
        for (RowData rd : dataRows) {
            Row outRow = outSheet.createRow(outRowIndex++);
            copyRow(rd.source, outRow);
        }

        String summary = String.format(
                "Sorted by '%s' %s %s → sheet '%s' (%d rows)",
                column, directionStr, numeric ? "(numeric)" : "(text)",
                outSheetName, dataRows.size());
        log.debug("SORT step {}: {}", step.getOrder(), summary);

        context.setActiveSheet(outSheetName);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .rowCount(dataRows.size())
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    private boolean isRowEmpty(Row row) {
        int lastCell = row.getLastCellNum();
        if (lastCell < 0) return true;
        for (int c = 0; c < lastCell; c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                if (cell.getCellType() == CellType.STRING) {
                    String v = cell.getStringCellValue();
                    if (v != null && !v.trim().isEmpty()) return false;
                } else {
                    return false;
                }
            }
        }
        return true;
    }

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

    /**
     * Holds cell values from a row for in-memory sorting.
     */
    private static class RowData {
        private final Row source;
        private final Map<Integer, String> stringValues = new HashMap<>();
        private final Map<Integer, Double> numericValues = new HashMap<>();
        private int lastCellNum;

        RowData(Row row) {
            this.source = row;
            this.lastCellNum = row.getLastCellNum();
            for (int c = 0; c < lastCellNum; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) {
                    stringValues.put(c, "");
                    numericValues.put(c, 0.0);
                } else {
                    stringValues.put(c, cellToString(cell));
                    if (cell.getCellType() == CellType.NUMERIC) {
                        numericValues.put(c, cell.getNumericCellValue());
                    } else {
                        numericValues.put(c, 0.0);
                    }
                }
            }
        }

        String getString(int col) {
            return stringValues.getOrDefault(col, "");
        }

        double getNumeric(int col) {
            return numericValues.getOrDefault(col, 0.0);
        }
    }

    private static String cellToString(Cell cell) {
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
}
