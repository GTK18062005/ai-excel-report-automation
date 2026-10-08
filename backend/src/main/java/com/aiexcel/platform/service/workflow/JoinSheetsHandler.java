package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * Handler for {@code JOIN_SHEETS}.
 * <p>
 * Joins two sheets using a shared column. Both sheets are read from the context's
 * available sheets (input or output). Supports INNER and LEFT join types.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code leftSheet} — required: name of the left-side sheet</li>
 *   <li>{@code rightSheet} — required: name of the right-side sheet</li>
 *   <li>{@code joinColumn} — required: column name present in both sheets</li>
 *   <li>{@code joinType} — required: {@code INNER} or {@code LEFT}</li>
 * </ul>
 * <p>
 * Duplicate join keys in the right sheet produce one output row per match
 * (cartesian product on that key). Duplicate join keys in the left sheet are
 * processed in the same manner. Null or blank join-key values are skipped.
 * <p>
 * Column naming: left-sheet columns use their original names; right-sheet
 * columns that are not the join column are prefixed with {@code "right."}.
 * If a prefixed name still conflicts, the right-hand copy of the join column
 * is omitted from output (the left join key is used).
 */
@Component
@Slf4j
public class JoinSheetsHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "JOIN_SHEETS";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Map<String, String> params = step.getParameters();
        if (params == null) {
            throw new WorkflowExecutionException(
                    "JOIN_SHEETS step " + step.getOrder() + ": parameters must be provided");
        }

        String leftSheetName  = getRequired(params, "leftSheet", step.getOrder());
        String rightSheetName = getRequired(params, "rightSheet", step.getOrder());
        String joinColumn     = getRequired(params, "joinColumn", step.getOrder());
        String joinTypeStr    = getRequired(params, "joinType", step.getOrder()).trim().toUpperCase();

        if (!SUPPORTED_JOIN_TYPES.contains(joinTypeStr)) {
            throw new WorkflowExecutionException(
                    "JOIN_SHEETS step " + step.getOrder()
                    + ": parameter 'joinType' must be one of " + SUPPORTED_JOIN_TYPES
                    + ", got '" + joinTypeStr + "'");
        }

        Sheet leftSheet  = ExecutionUtils.findSheet(context, leftSheetName, step.getOrder());
        Sheet rightSheet = ExecutionUtils.findSheet(context, rightSheetName, step.getOrder());

        // ── Read headers ───────────────────────────────────────────────────────
        Row leftHeader  = getHeaderRow(leftSheet, step.getOrder(), leftSheetName);
        Row rightHeader = getHeaderRow(rightSheet, step.getOrder(), rightSheetName);

        List<String> leftCols  = readHeaders(leftHeader);
        List<String> rightCols = readHeaders(rightHeader);

        int leftJoinIdx  = leftCols.indexOf(joinColumn);
        int rightJoinIdx = rightCols.indexOf(joinColumn);

        if (leftJoinIdx < 0) {
            throw new WorkflowExecutionException(
                    "JOIN_SHEETS step " + step.getOrder()
                    + ": join column '" + joinColumn + "' not found in left sheet '" + leftSheetName + "'");
        }
        if (rightJoinIdx < 0) {
            throw new WorkflowExecutionException(
                    "JOIN_SHEETS step " + step.getOrder()
                    + ": join column '" + joinColumn + "' not found in right sheet '" + rightSheetName + "'");
        }

        // ── Read data rows into maps keyed by join value ──────────────────────
        Map<String, List<String[]>> rightMap = buildKeyMap(rightSheet, rightCols, rightJoinIdx);
        Map<String, List<String[]>> leftMap  = buildKeyMap(leftSheet,  leftCols,  leftJoinIdx);

        // ── Build output column list ──────────────────────────────────────────
        // Left columns (all), then right columns except the join column (prefixed)
        List<String> outHeaders = new ArrayList<>(leftCols);
        Set<String> usedNames = new HashSet<>(outHeaders);
        for (String rc : rightCols) {
            if (rc.equals(joinColumn)) continue;
            String prefixed = prefixIfDuplicate(rc, usedNames, "right.");
            outHeaders.add(prefixed);
        }

        // ── Create output sheet ───────────────────────────────────────────────
        String outSheetName = leftSheetName + "_joined";
        int suffix = 1;
        while (context.getOutputWorkbook().getSheet(outSheetName) != null) {
            outSheetName = leftSheetName + "_joined_" + (++suffix);
        }
        Sheet outSheet = context.getOutputWorkbook().createSheet(outSheetName);

        // Write header
        Row outHeader = outSheet.createRow(0);
        for (int c = 0; c < outHeaders.size(); c++) {
            outHeader.createCell(c).setCellValue(outHeaders.get(c));
        }

        int outRowIdx = 1;
        int matchedRows = 0;
        int unmatchedLeftRows = 0;

        // Determine which left keys to process (union for LEFT, intersection for INNER)
        Set<String> leftKeys = new LinkedHashSet<>(leftMap.keySet());

        for (String leftKey : leftKeys) {
            List<String[]> leftRows  = leftMap.get(leftKey);
            List<String[]> rightRows = rightMap.get(leftKey);

            if (rightRows != null && !rightRows.isEmpty()) {
                // Key found in both sheets — emit one row per combination
                for (String[] lr : leftRows) {
                    for (String[] rr : rightRows) {
                        writeJoinedRow(outSheet, outRowIdx++, lr, leftCols.size(),
                                       rr, rightCols, rightJoinIdx, joinColumn);
                        matchedRows++;
                    }
                }
            } else if ("LEFT".equals(joinTypeStr)) {
                // No match — emit left row with blanks for right side
                for (String[] lr : leftRows) {
                    writeLeftOnlyRow(outSheet, outRowIdx++, lr, leftCols.size(), rightCols.size());
                    unmatchedLeftRows++;
                }
            }
            // INNER: unmatched left rows are silently dropped
        }

        context.setActiveSheet(outSheetName);

        String summary = String.format(
                "JOIN %s [%s=%s] → sheet '%s' (%d matched, %d unmatched left)",
                joinTypeStr, leftSheetName, rightSheetName,
                outSheetName, matchedRows, unmatchedLeftRows);
        log.debug("JOIN_SHEETS step {}: {}", step.getOrder(), summary);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .columns(outHeaders)
                .rowCount(matchedRows + unmatchedLeftRows)
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private String getRequired(Map<String, String> params, String key, int stepOrder) {
        String val = params.get(key);
        if (val == null || val.isBlank()) {
            throw new WorkflowExecutionException(
                    "JOIN_SHEETS step " + stepOrder + ": parameter '" + key + "' is required");
        }
        return val;
    }

    private Row getHeaderRow(Sheet sheet, int stepOrder, String sheetName) {
        int first = sheet.getFirstRowNum();
        if (first < 0) {
            throw new WorkflowExecutionException(
                    "JOIN_SHEETS step " + stepOrder + ": sheet '" + sheetName + "' is empty");
        }
        Row header = sheet.getRow(first);
        if (header == null) {
            throw new WorkflowExecutionException(
                    "JOIN_SHEETS step " + stepOrder + ": sheet '" + sheetName + "' has no header row");
        }
        return header;
    }

    private List<String> readHeaders(Row header) {
        List<String> cols = new ArrayList<>();
        int count = header.getLastCellNum();
        for (int c = 0; c < count; c++) {
            Cell cell = header.getCell(c);
            cols.add(cellToString(cell).trim());
        }
        return cols;
    }

    /**
     * Reads all data rows (skipping header) and groups them by the join column value.
     * Rows with null/blank join keys are skipped.
     * Duplicate keys produce a list with multiple entries.
     */
    private Map<String, List<String[]>> buildKeyMap(Sheet sheet, List<String> cols, int joinIdx) {
        Map<String, List<String[]>> map = new LinkedHashMap<>();
        int first = sheet.getFirstRowNum();
        int last  = sheet.getLastRowNum();
        for (int r = first + 1; r <= last; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String[] values = rowToStringArray(row, cols.size());
            String key = values[joinIdx];
            if (key == null || key.isBlank()) continue; // skip null/blank keys
            map.computeIfAbsent(key, k -> new ArrayList<>()).add(values);
        }
        return map;
    }

    private String[] rowToStringArray(Row row, int colCount) {
        String[] values = new String[colCount];
        for (int c = 0; c < colCount; c++) {
            Cell cell = row.getCell(c);
            values[c] = cellToString(cell);
        }
        return values;
    }

    /**
     * Writes a fully-joined output row: left values + right values (right join column excluded).
     */
    private void writeJoinedRow(Sheet outSheet, int outRowIdx,
                                 String[] leftVals, int leftCount,
                                 String[] rightVals, List<String> rightCols, int rightJoinIdx,
                                 String joinColumn) {
        Row outRow = outSheet.createRow(outRowIdx);
        int outCol = 0;

        // Left values
        for (int c = 0; c < leftCount; c++) {
            writeCell(outRow, outCol++, leftVals[c]);
        }

        // Right values, skipping the join column
        for (int c = 0; c < rightVals.length; c++) {
            if (c == rightJoinIdx) continue;
            writeCell(outRow, outCol++, rightVals[c]);
        }
    }

    /**
     * Writes a left-only row: left values + blank cells for right side.
     */
    private void writeLeftOnlyRow(Sheet outSheet, int outRowIdx,
                                   String[] leftVals, int leftCount, int rightNonJoinCount) {
        Row outRow = outSheet.createRow(outRowIdx);
        int outCol = 0;
        for (int c = 0; c < leftCount; c++) {
            writeCell(outRow, outCol++, leftVals[c]);
        }
        for (int c = 0; c < rightNonJoinCount; c++) {
            outRow.createCell(outCol++); // blank
        }
    }

    private void writeCell(Row row, int col, String value) {
        if (value == null || value.isEmpty()) {
            row.createCell(col);
            return;
        }
        // Try parse as number for better type fidelity
        try {
            double d = Double.parseDouble(value.trim());
            row.createCell(col).setCellValue(d);
        } catch (NumberFormatException e) {
            row.createCell(col).setCellValue(value);
        }
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
            case BLANK -> "";
            default -> "";
        };
    }

    /**
     * Returns the given name prefixed with the prefix if it already exists in usedNames.
     * Also adds the chosen name to usedNames.
     */
    private String prefixIfDuplicate(String name, Set<String> usedNames, String prefix) {
        if (!usedNames.contains(name)) {
            usedNames.add(name);
            return name;
        }
        String candidate = prefix + name;
        if (!usedNames.contains(candidate)) {
            usedNames.add(candidate);
            return candidate;
        }
        // Both original and prefixed collide — find a unique suffix
        int suffix = 2;
        while (usedNames.contains(candidate + "_" + suffix)) {
            suffix++;
        }
        String result = candidate + "_" + suffix;
        usedNames.add(result);
        return result;
    }

    private static final Set<String> SUPPORTED_JOIN_TYPES = Set.of("INNER", "LEFT");
}
