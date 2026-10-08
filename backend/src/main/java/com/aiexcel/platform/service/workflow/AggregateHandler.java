package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * Handler for {@code AGGREGATE}.
 * <p>
 * Groups rows by one or more columns and computes an aggregate function on a
 * numeric column, producing one row per unique group key combination.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code groupColumns} — optional: comma-separated list of columns to group by;
 *       if omitted, a single summary row is produced</li>
 *   <li>{@code aggregateColumn} — required: numeric column to aggregate</li>
 *   <li>{@code aggregateFunction} — required: one of SUM, COUNT, AVERAGE, MIN, MAX</li>
 * </ul>
 */
@Component
@Slf4j
public class AggregateHandler implements WorkflowOperationHandler {

    private static final Set<String> ALLOWED_FUNCTIONS = Set.of("SUM", "COUNT", "AVERAGE", "MIN", "MAX");

    @Override
    public String getSupportedOperation() {
        return "AGGREGATE";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        Map<String, String> params = step.getParameters();
        if (params == null) {
            throw new WorkflowExecutionException(
                    "AGGREGATE step " + step.getOrder() + ": parameters must be provided");
        }

        String aggregateColumn = params.get("aggregateColumn");
        if (aggregateColumn == null || aggregateColumn.isBlank()) {
            throw new WorkflowExecutionException(
                    "AGGREGATE step " + step.getOrder() + ": parameter 'aggregateColumn' is required");
        }

        String functionStr = params.get("aggregateFunction");
        if (functionStr == null || functionStr.isBlank()) {
            throw new WorkflowExecutionException(
                    "AGGREGATE step " + step.getOrder() + ": parameter 'aggregateFunction' is required");
        }
        functionStr = functionStr.trim().toUpperCase();
        if (!ALLOWED_FUNCTIONS.contains(functionStr)) {
            throw new WorkflowExecutionException(
                    "AGGREGATE step " + step.getOrder() + ": unsupported function '" + functionStr
                            + "'. Supported: " + String.join(", ", ALLOWED_FUNCTIONS));
        }

        String groupColumnsRaw = params.get("groupColumns");
        List<String> groupColumns = new ArrayList<>();
        if (groupColumnsRaw != null && !groupColumnsRaw.isBlank()) {
            for (String col : groupColumnsRaw.split(",")) {
                String trimmed = col.trim();
                if (!trimmed.isEmpty()) groupColumns.add(trimmed);
            }
        }

        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "AGGREGATE step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' is empty");
        }
        Row header = sourceSheet.getRow(firstRow);
        if (header == null) {
            throw new WorkflowExecutionException(
                    "AGGREGATE step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' has no header row");
        }

        int colCount = header.getLastCellNum();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < colCount; c++) {
            Cell hc = header.getCell(c);
            headers.add(hc != null && hc.getCellType() == CellType.STRING
                    ? hc.getStringCellValue().trim() : "");
        }

        // Validate aggregate column
        int aggColIdx = headers.indexOf(aggregateColumn);
        if (aggColIdx < 0) {
            throw new WorkflowExecutionException(
                    "AGGREGATE step " + step.getOrder() + ": aggregate column '" + aggregateColumn
                            + "' not found in sheet '" + sourceSheetName + "'");
        }

        // Validate group columns
        List<Integer> groupColIndices = new ArrayList<>();
        for (String gc : groupColumns) {
            int idx = headers.indexOf(gc);
            if (idx < 0) {
                throw new WorkflowExecutionException(
                        "AGGREGATE step " + step.getOrder() + ": group column '" + gc
                                + "' not found in sheet '" + sourceSheetName + "'");
            }
            groupColIndices.add(idx);
        }

        // Collect rows
        int lastRow = sourceSheet.getLastRowNum();
        List<double[]> dataRows = new ArrayList<>();
        List<String[]> groupKeys = new ArrayList<>();

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row row = sourceSheet.getRow(r);
            if (row == null) continue;
            if (isRowEmpty(row)) continue;

            String[] gk = new String[groupColumns.size()];
            for (int i = 0; i < groupColumns.size(); i++) {
                Cell c = row.getCell(groupColIndices.get(i));
                gk[i] = cellToString(c);
            }

            Cell aggCell = row.getCell(aggColIdx);
            if (aggCell != null && aggCell.getCellType() == CellType.NUMERIC) {
                groupKeys.add(gk);
                dataRows.add(new double[]{aggCell.getNumericCellValue()});
            }
        }

        if (dataRows.isEmpty()) {
            throw new WorkflowExecutionException(
                    "AGGREGATE step " + step.getOrder() + ": no numeric values found in column '"
                            + aggregateColumn + "'");
        }

        // Compute aggregates per group
        Map<String, List<Double>> groups = new LinkedHashMap<>();
        for (int i = 0; i < groupKeys.size(); i++) {
            String key = String.join("§", groupKeys.get(i));
            groups.computeIfAbsent(key, k -> new ArrayList<>())
                    .add(dataRows.get(i)[0]);
        }

        String summaryColumn = aggregateColumn + "_" + functionStr;

        // Build output sheet
        String outSheetName = sourceSheetName + "_agg";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_agg_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Header: group columns + summary column
        Row outHeader = outSheet.createRow(0);
        int outCol = 0;
        for (String gc : groupColumns) {
            outHeader.createCell(outCol++).setCellValue(gc);
        }
        outHeader.createCell(outCol).setCellValue(summaryColumn);

        int outRowIdx = 1;
        for (Map.Entry<String, List<Double>> entry : groups.entrySet()) {
            List<Double> vals = entry.getValue();
            double result = computeAggregate(functionStr, vals);

            Row outRow = outSheet.createRow(outRowIdx++);
            int writeCol = 0;

            if (!groupColumns.isEmpty()) {
                String[] parts = entry.getKey().split("§");
                for (String part : parts) {
                    outRow.createCell(writeCol++).setCellValue(part);
                }
            }

            Cell aggCell = outRow.createCell(writeCol);
            aggCell.setCellValue(result);
        }

        String summary = String.format(
                "%s(%s) → sheet '%s' (%d groups)",
                functionStr, aggregateColumn, outSheetName, groups.size());
        log.debug("AGGREGATE step {}: {}", step.getOrder(), summary);

        context.setActiveSheet(outSheetName);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .rowCount(groups.size())
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    private double computeAggregate(String function, List<Double> values) {
        return switch (function) {
            case "SUM" -> values.stream().mapToDouble(Double::doubleValue).sum();
            case "COUNT" -> values.size();
            case "AVERAGE" -> values.stream().mapToDouble(Double::doubleValue).average().orElse(0.0);
            case "MIN" -> values.stream().mapToDouble(Double::doubleValue).min().orElse(0.0);
            case "MAX" -> values.stream().mapToDouble(Double::doubleValue).max().orElse(0.0);
            default -> throw new IllegalStateException("Unknown function: " + function);
        };
    }

    private boolean isRowEmpty(Row row) {
        int lastCell = row.getLastCellNum();
        if (lastCell < 0) return true;
        for (int c = 0; c < lastCell; c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK) return false;
        }
        return true;
    }

    private String cellToString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> {
                String v = cell.getStringCellValue();
                yield v != null ? v.trim() : "";
            }
            case NUMERIC -> {
                double d = cell.getNumericCellValue();
                yield (d == Math.floor(d) && !Double.isInfinite(d))
                        ? String.valueOf((long) d) : String.valueOf(d);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }
}
