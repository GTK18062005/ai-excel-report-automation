package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * Handler for {@code GROUP_BY}.
 * <p>
 * Groups rows by one or more columns and produces a new sheet with one row
 * per unique combination of group keys. All non-group columns are preserved
 * as-is from the first row encountered for each group.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code groupColumns} — required: comma-separated list of column names to group by</li>
 * </ul>
 */
@Component
@Slf4j
public class GroupByHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "GROUP_BY";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        Map<String, String> params = step.getParameters();
        if (params == null) {
            throw new WorkflowExecutionException(
                    "GROUP_BY step " + step.getOrder() + ": parameters must be provided");
        }

        String groupColumnsRaw = params.get("groupColumns");
        if (groupColumnsRaw == null || groupColumnsRaw.isBlank()) {
            throw new WorkflowExecutionException(
                    "GROUP_BY step " + step.getOrder() + ": parameter 'groupColumns' is required");
        }

        List<String> groupColumns = new ArrayList<>();
        for (String col : groupColumnsRaw.split(",")) {
            String trimmed = col.trim();
            if (!trimmed.isEmpty()) groupColumns.add(trimmed);
        }
        if (groupColumns.isEmpty()) {
            throw new WorkflowExecutionException(
                    "GROUP_BY step " + step.getOrder() + ": at least one group column is required");
        }

        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "GROUP_BY step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' is empty");
        }
        Row header = sourceSheet.getRow(firstRow);
        if (header == null) {
            throw new WorkflowExecutionException(
                    "GROUP_BY step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' has no header row");
        }

        int colCount = header.getLastCellNum();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < colCount; c++) {
            Cell hc = header.getCell(c);
            headers.add(hc != null && hc.getCellType() == CellType.STRING
                    ? hc.getStringCellValue().trim() : "");
        }

        // Validate group columns
        List<Integer> groupColIndices = new ArrayList<>();
        for (String groupCol : groupColumns) {
            int idx = headers.indexOf(groupCol);
            if (idx < 0) {
                throw new WorkflowExecutionException(
                        "GROUP_BY step " + step.getOrder() + ": group column '" + groupCol
                                + "' not found in sheet '" + sourceSheetName + "'");
            }
            groupColIndices.add(idx);
        }

        // Collect data rows and group them
        int lastRow = sourceSheet.getLastRowNum();
        Map<String, RowData> groupMap = new LinkedHashMap<>();

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row row = sourceSheet.getRow(r);
            if (row == null) continue;

            String key = buildGroupKey(row, groupColIndices);
            if (!groupMap.containsKey(key)) {
                groupMap.put(key, new RowData(row, colCount));
            }
        }

        // Create output sheet
        String outSheetName = sourceSheetName + "_grouped";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_grouped_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Write header (all original columns)
        Row outHeader = outSheet.createRow(0);
        for (int c = 0; c < colCount; c++) {
            Cell hc = header.getCell(c);
            if (hc != null) {
                Cell nc = outHeader.createCell(c);
                copyCellValue(hc, nc);
            }
        }

        // Write grouped rows
        int outRowIdx = 1;
        for (Map.Entry<String, RowData> entry : groupMap.entrySet()) {
            RowData rd = entry.getValue();
            Row outRow = outSheet.createRow(outRowIdx++);
            for (int c = 0; c < rd.values.length; c++) {
                if (rd.values[c] != null) {
                    Cell nc = outRow.createCell(c);
                    nc.setCellValue(rd.values[c]);
                }
            }
        }

        String summary = String.format(
                "Grouped by %s → sheet '%s' (%d unique groups from %d rows)",
                groupColumns, outSheetName, groupMap.size(), lastRow - firstRow);
        log.debug("GROUP_BY step {}: {}", step.getOrder(), summary);

        context.setActiveSheet(outSheetName);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .rowCount(groupMap.size())
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    private String buildGroupKey(Row row, List<Integer> groupColIndices) {
        StringBuilder sb = new StringBuilder();
        for (int idx : groupColIndices) {
            if (sb.length() > 0) sb.append("§");
            Cell cell = row.getCell(idx);
            sb.append(cellToKeyString(cell));
        }
        return sb.toString();
    }

    private String cellToKeyString(Cell cell) {
        if (cell == null) return "∅";
        return switch (cell.getCellType()) {
            case STRING -> {
                String v = cell.getStringCellValue();
                yield v != null ? v.trim() : "";
            }
            case NUMERIC -> {
                double d = cell.getNumericCellValue();
                if (d == Math.floor(d) && !Double.isInfinite(d)) yield String.valueOf((long) d);
                yield String.valueOf(d);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case BLANK -> "";
            default -> "";
        };
    }

    private void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING -> target.setCellValue(source.getStringCellValue());
            case NUMERIC -> target.setCellValue(source.getNumericCellValue());
            case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
            case BLANK -> target.setCellType(CellType.BLANK);
            case FORMULA -> {
                try { target.setCellFormula(source.getCellFormula()); }
                catch (Exception ignored) { target.setCellType(CellType.BLANK); }
            }
            default -> {}
        }
    }

    /** Holds cell values for one row as strings (for grouping key) and numerics (for aggregate). */
    private static class RowData {
        private final String[] values;

        RowData(Row source, int colCount) {
            this.values = new String[colCount];
            for (int c = 0; c < colCount; c++) {
                Cell cell = source.getCell(c);
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING -> {
                            String v = cell.getStringCellValue();
                            values[c] = v != null ? v.trim() : "";
                        }
                        case NUMERIC -> {
                            double d = cell.getNumericCellValue();
                            values[c] = (d == Math.floor(d) && !Double.isInfinite(d))
                                    ? String.valueOf((long) d) : String.valueOf(d);
                        }
                        case BOOLEAN -> values[c] = String.valueOf(cell.getBooleanCellValue());
                    }
                }
            }
        }
    }
}
