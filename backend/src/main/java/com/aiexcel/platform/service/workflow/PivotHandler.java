package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * Handler for {@code PIVOT}.
 * <p>
 * Produces a pivot table from the source sheet. Each unique value in the pivot
 * column becomes a column header in the output. Each unique value in the row
 * column becomes a row label. Cells are populated by aggregating the value column.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code sourceSheet} — required: sheet to pivot; defaults to active sheet if omitted</li>
 *   <li>{@code rowColumn} — required: column whose unique values form output rows</li>
 *   <li>{@code pivotColumn} — required: column whose unique values form output column headers</li>
 *   <li>{@code valueColumn} — required: numeric column whose values are aggregated</li>
 *   <li>{@code aggregationFunction} — required: one of SUM, COUNT, AVERAGE, MIN, MAX</li>
 * </ul>
 * <p>
 * Empty input or a pivot column with only one value produces a sensible output
 * with a single data column. Rows with blank pivot or row keys are skipped.
 * The output sheet name is suffixed with {@code "_pivot"}.
 */
@Component
@Slf4j
public class PivotHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "PIVOT";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        Map<String, String> params = step.getParameters();
        if (params == null) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + step.getOrder() + ": parameters must be provided");
        }

        String sourceSheetName = params.containsKey("sourceSheet")
                ? params.get("sourceSheet")
                : ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        String rowColumn    = getRequired(params, "rowColumn",    step.getOrder());
        String pivotColumn = getRequired(params, "pivotColumn", step.getOrder());
        String valueColumn = getRequired(params, "valueColumn", step.getOrder());
        String funcStr     = getRequired(params, "aggregationFunction", step.getOrder()).trim().toUpperCase();

        if (!ALLOWED_FUNCTIONS.contains(funcStr)) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + step.getOrder()
                    + ": parameter 'aggregationFunction' must be one of " + ALLOWED_FUNCTIONS
                    + ", got '" + funcStr + "'");
        }

        // ── Read headers ───────────────────────────────────────────────────────
        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + step.getOrder() + ": source sheet is empty");
        }
        Row headerRow = sourceSheet.getRow(firstRow);
        if (headerRow == null) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + step.getOrder() + ": source sheet has no header row");
        }

        int colCount = headerRow.getLastCellNum();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < colCount; c++) {
            Cell hc = headerRow.getCell(c);
            headers.add(hc != null && hc.getCellType() == CellType.STRING
                    ? hc.getStringCellValue().trim() : "");
        }

        int rowIdx    = headers.indexOf(rowColumn);
        int pivotIdx  = headers.indexOf(pivotColumn);
        int valueIdx  = headers.indexOf(valueColumn);

        if (rowIdx < 0) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + step.getOrder()
                    + ": row column '" + rowColumn + "' not found in sheet '" + sourceSheetName + "'");
        }
        if (pivotIdx < 0) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + step.getOrder()
                    + ": pivot column '" + pivotColumn + "' not found in sheet '" + sourceSheetName + "'");
        }
        if (valueIdx < 0) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + step.getOrder()
                    + ": value column '" + valueColumn + "' not found in sheet '" + sourceSheetName + "'");
        }

        // ── Collect data and group by (row key, pivot key) ───────────────────
        int lastRow = sourceSheet.getLastRowNum();

        // rowKey → (pivotKey → list of values)
        Map<String, Map<String, List<Double>>> data = new LinkedHashMap<>();
        Set<String> pivotKeys = new LinkedHashSet<>();

        int dataRowCount = 0;

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row row = sourceSheet.getRow(r);
            if (row == null) continue;

            Cell rowCell    = row.getCell(rowIdx);
            Cell pivotCell = row.getCell(pivotIdx);
            Cell valueCell = row.getCell(valueIdx);

            String rowKey    = cellToString(rowCell).trim();
            String pivotKey  = cellToString(pivotCell).trim();

            if (rowKey.isEmpty() || pivotKey.isEmpty()) continue; // skip blank keys

            if (!data.containsKey(rowKey)) {
                data.put(rowKey, new LinkedHashMap<>());
            }

            Map<String, List<Double>> rowMap = data.get(rowKey);
            if (!rowMap.containsKey(pivotKey)) {
                rowMap.put(pivotKey, new ArrayList<>());
                pivotKeys.add(pivotKey);
            }

            if (valueCell != null && valueCell.getCellType() == CellType.NUMERIC) {
                rowMap.get(pivotKey).add(valueCell.getNumericCellValue());
                dataRowCount++;
            }
        }

        if (data.isEmpty()) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + step.getOrder()
                    + ": no valid data rows found for the specified columns");
        }

        // ── Build output ─────────────────────────────────────────────────────
        String outSheetName = sourceSheetName + "_pivot";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_pivot_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Header: rowColumn label + one column per pivot key (sorted for determinism)
        List<String> sortedPivotKeys = new ArrayList<>(pivotKeys);
        Collections.sort(sortedPivotKeys);

        Row outHeader = outSheet.createRow(0);
        outHeader.createCell(0).setCellValue(rowColumn);
        for (int c = 0; c < sortedPivotKeys.size(); c++) {
            outHeader.createCell(c + 1).setCellValue(sortedPivotKeys.get(c));
        }

        int outRowIdx = 1;
        List<String> sortedRowKeys = new ArrayList<>(data.keySet());
        Collections.sort(sortedRowKeys);

        for (String rowKey : sortedRowKeys) {
            Row outRow = outSheet.createRow(outRowIdx++);
            outRow.createCell(0).setCellValue(rowKey);

            Map<String, List<Double>> rowMap = data.get(rowKey);
            for (int c = 0; c < sortedPivotKeys.size(); c++) {
                String pivotKey = sortedPivotKeys.get(c);
                List<Double> vals = rowMap.get(pivotKey);
                Cell cell = outRow.createCell(c + 1);
                if (vals == null || vals.isEmpty()) {
                    // leave blank — no data for this combination
                } else {
                    double result = computeAggregate(funcStr, vals);
                    cell.setCellValue(result);
                }
            }
        }

        context.setActiveSheet(outSheetName);

        int outputRows = sortedRowKeys.size();
        String summary = String.format(
                "PIVOT %s x %s (%s) → sheet '%s' (%d row groups, %d pivot columns)",
                rowColumn, pivotColumn, funcStr, outSheetName, outputRows, sortedPivotKeys.size());
        log.debug("PIVOT step {}: {}", step.getOrder(), summary);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .columns(concat(
                        List.of(rowColumn),
                        new ArrayList<>(sortedPivotKeys)))
                .rowCount(outputRows)
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private String getRequired(Map<String, String> params, String key, int stepOrder) {
        String val = params.get(key);
        if (val == null || val.isBlank()) {
            throw new WorkflowExecutionException(
                    "PIVOT step " + stepOrder + ": parameter '" + key + "' is required");
        }
        return val;
    }

    private double computeAggregate(String function, List<Double> values) {
        return switch (function) {
            case "SUM"     -> values.stream().mapToDouble(Double::doubleValue).sum();
            case "COUNT"   -> values.size();
            case "AVERAGE" -> values.stream().mapToDouble(Double::doubleValue).average().orElse(0.0);
            case "MIN"     -> values.stream().mapToDouble(Double::doubleValue).min().orElse(0.0);
            case "MAX"     -> values.stream().mapToDouble(Double::doubleValue).max().orElse(0.0);
            default -> throw new IllegalStateException("Unknown function: " + function);
        };
    }

    private String cellToString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING  -> cell.getStringCellValue() != null ? cell.getStringCellValue().trim() : "";
            case NUMERIC -> {
                double d = cell.getNumericCellValue();
                yield (d == Math.floor(d) && !Double.isInfinite(d))
                        ? String.valueOf((long) d) : String.valueOf(d);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    private static <T> List<T> concat(List<T> first, List<T> second) {
        List<T> result = new ArrayList<>(first);
        result.addAll(second);
        return result;
    }

    private static final Set<String> ALLOWED_FUNCTIONS = Set.of("SUM", "COUNT", "AVERAGE", "MIN", "MAX");
}
