package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import java.util.HashSet;
import java.util.Set;

/**
 * Handler for {@code REMOVE_DUPLICATES}.
 * <p>
 * Copies rows from the source sheet to a new output sheet, dropping
 * rows whose concatenated-cell string matches a previously-seen row.
 * If {@code inputColumns} is provided, the de-duplication key is
 * the concatenation of values in those columns only; otherwise the
 * full row is used as the key.
 * The header row is always kept.
 */
@Component
@Slf4j
public class RemoveDuplicatesHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "REMOVE_DUPLICATES";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "REMOVE_DUPLICATES step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' is empty");
        }

        Row header = sourceSheet.getRow(firstRow);
        if (header == null) {
            throw new WorkflowExecutionException(
                    "REMOVE_DUPLICATES step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' has no header row");
        }

        // Determine which column indices to use as the de-dup key
        int[] keyIndices;
        if (step.getInputColumns() != null && !step.getInputColumns().isEmpty()) {
            int colCount = header.getLastCellNum();
            java.util.List<String> sourceHeaders = new java.util.ArrayList<>();
            for (int c = 0; c < colCount; c++) {
                Cell hc = header.getCell(c);
                sourceHeaders.add(hc != null ? cellToTrimmedString(hc) : "");
            }
            keyIndices = new int[step.getInputColumns().size()];
            int i = 0;
            for (String requested : step.getInputColumns()) {
                int idx = sourceHeaders.indexOf(requested);
                if (idx < 0) {
                    throw new WorkflowExecutionException(
                            "REMOVE_DUPLICATES step " + step.getOrder() + ": column '" + requested
                                    + "' not found in source sheet '" + sourceSheetName + "'");
                }
                keyIndices[i++] = idx;
            }
        } else {
            int colCount = header.getLastCellNum();
            keyIndices = new int[colCount];
            for (int i = 0; i < colCount; i++) keyIndices[i] = i;
        }

        // Create output sheet
        String outSheetName = sourceSheetName + "_unique";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_unique_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Copy header
        copyRow(header, outSheet.createRow(0));

        Set<String> seen = new HashSet<>();
        int rowsKept = 0;
        int rowsRemoved = 0;
        int outRowIndex = 1;
        int lastRow = sourceSheet.getLastRowNum();

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row sourceRow = sourceSheet.getRow(r);
            if (sourceRow == null) continue;
            String key = rowKey(sourceRow, keyIndices);
            if (!seen.add(key)) {
                rowsRemoved++;
                continue;
            }
            Row outRow = outSheet.createRow(outRowIndex++);
            copyRow(sourceRow, outRow);
            rowsKept++;
        }

        String summary = String.format(
                "Removed %d duplicate rows → sheet '%s' (%d rows kept)",
                rowsRemoved, outSheetName, rowsKept);
        log.debug("REMOVE_DUPLICATES step {}: {}", step.getOrder(), summary);

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

    private String rowKey(Row row, int[] indices) {
        StringBuilder sb = new StringBuilder();
        for (int idx : indices) {
            Cell c = row.getCell(idx);
            sb.append("§").append(c == null ? "" : cellToTrimmedString(c));
        }
        return sb.toString();
    }

    private String cellToTrimmedString(Cell cell) {
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
            case BLANK -> "";
            default -> "";
        };
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
            case BLANK -> target.setCellType(org.apache.poi.ss.usermodel.CellType.BLANK);
            case FORMULA -> {
                try {
                    target.setCellFormula(source.getCellFormula());
                } catch (Exception ignored) {
                    target.setCellType(org.apache.poi.ss.usermodel.CellType.BLANK);
                }
            }
            default -> { /* unsupported, leave blank */ }
        }
    }
}
