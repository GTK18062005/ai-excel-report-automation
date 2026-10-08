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

import java.util.List;

/**
 * Handler for {@code REMOVE_EMPTY_ROWS}.
 * <p>
 * Copies all rows from the source sheet to a new output sheet,
 * skipping rows where every cell is empty or null.
 * The header row (first row) is always kept.
 */
@Component
@Slf4j
public class RemoveEmptyRowsHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "REMOVE_EMPTY_ROWS";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        String outSheetName = sourceSheetName + "_clean";
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_clean_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        int firstRow = sourceSheet.getFirstRowNum();
        int lastRow = sourceSheet.getLastRowNum();

        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "REMOVE_EMPTY_ROWS step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' is empty");
        }

        // Copy header row (always keep)
        Row header = sourceSheet.getRow(firstRow);
        if (header != null) {
            copyRow(header, outSheet.createRow(0));
        }

        int rowsKept = 0;
        int rowsRemoved = 0;
        int outRowIndex = 1;

        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row sourceRow = sourceSheet.getRow(r);
            if (sourceRow == null) {
                rowsRemoved++;
                continue;
            }
            if (isRowEmpty(sourceRow)) {
                rowsRemoved++;
                continue;
            }
            Row outRow = outSheet.createRow(outRowIndex++);
            copyRow(sourceRow, outRow);
            rowsKept++;
        }

        String summary = String.format(
                "Removed %d empty rows → sheet '%s' (%d rows kept)",
                rowsRemoved, outSheetName, rowsKept);
        log.debug("REMOVE_EMPTY_ROWS step {}: {}", step.getOrder(), summary);

        context.setActiveSheet(outSheetName);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .columns(List.of())
                .rowCount(rowsKept)
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    private boolean isRowEmpty(Row row) {
        int lastCell = row.getLastCellNum();
        if (lastCell < 0) return true;
        for (int c = 0; c < lastCell; c++) {
            Cell cell = row.getCell(c);
            if (!isCellEmpty(cell)) return false;
        }
        return true;
    }

    private boolean isCellEmpty(Cell cell) {
        if (cell == null) return true;
        CellType ct = cell.getCellType();
        if (ct == CellType.BLANK) return true;
        if (ct == CellType.STRING) {
            String v = cell.getStringCellValue();
            return v == null || v.trim().isEmpty();
        }
        return false;
    }

    private void copyRow(Row source, Row target) {
        int lastCell = source.getLastCellNum();
        for (int c = 0; c < lastCell; c++) {
            Cell srcCell = source.getCell(c);
            if (srcCell == null) continue;
            Cell tgtCell = target.createCell(c);
            copyCellValue(srcCell, tgtCell);
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
