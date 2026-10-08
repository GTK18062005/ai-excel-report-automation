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

import java.util.ArrayList;
import java.util.List;

/**
 * Handler for {@code SELECT_COLUMNS}.
 * <p>
 * Produces a new sheet in the output workbook that contains only the
 * columns listed in {@code inputColumns}, in that order.
 * The source data comes from the input workbook (read-only).
 */
@Component
@Slf4j
public class SelectColumnsHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "SELECT_COLUMNS";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook output = context.getOutputWorkbook();

        if (step.getInputColumns() == null || step.getInputColumns().isEmpty()) {
            throw new WorkflowExecutionException(
                    "SELECT_COLUMNS step " + step.getOrder() + ": at least one input column must be specified");
        }

        String sourceSheetName = ExecutionUtils.resolveSourceSheet(step, context);
        Sheet sourceSheet = ExecutionUtils.findSheet(context, sourceSheetName, step.getOrder());

        // Read header from source
        int firstRow = sourceSheet.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "SELECT_COLUMNS step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' is empty");
        }
        Row headerRow = sourceSheet.getRow(firstRow);
        if (headerRow == null) {
            throw new WorkflowExecutionException(
                    "SELECT_COLUMNS step " + step.getOrder() + ": source sheet '" + sourceSheetName + "' has no header row");
        }

        List<String> sourceHeaders = new ArrayList<>();
        for (int c = 0; c < headerRow.getLastCellNum(); c++) {
            Cell cell = headerRow.getCell(c);
            sourceHeaders.add(cell != null && cell.getCellType() == CellType.STRING
                    ? cell.getStringCellValue().trim() : "");
        }

        // Resolve indices of requested columns in order
        List<Integer> selectedIndices = new ArrayList<>();
        for (String requestedCol : step.getInputColumns()) {
            if (requestedCol == null || requestedCol.isBlank()) continue;
            int idx = sourceHeaders.indexOf(requestedCol);
            if (idx < 0) {
                throw new WorkflowExecutionException(
                        "SELECT_COLUMNS step " + step.getOrder() + ": column '" + requestedCol
                                + "' not found in source sheet '" + sourceSheetName + "'");
            }
            selectedIndices.add(idx);
        }

        // Create output sheet
        String outSheetName = sourceSheetName + "_selected";
        // POI disallows duplicates — ensure unique
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = sourceSheetName + "_selected_" + (++suffix);
        }
        Sheet outSheet = output.createSheet(outSheetName);

        // Write header
        Row outHeader = outSheet.createRow(0);
        for (int c = 0; c < step.getInputColumns().size(); c++) {
            outHeader.createCell(c).setCellValue(step.getInputColumns().get(c));
        }

        // Copy data rows
        int lastRow = sourceSheet.getLastRowNum();
        int dataRows = 0;
        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row sourceDataRow = sourceSheet.getRow(r);
            if (sourceDataRow == null) continue;
            Row outRow = outSheet.createRow(outSheet.getLastRowNum() + 1);
            for (int c = 0; c < selectedIndices.size(); c++) {
                Cell sourceCell = sourceDataRow.getCell(selectedIndices.get(c));
                Cell newCell = outRow.createCell(c);
                if (sourceCell != null) {
                    copyCellValue(sourceCell, newCell);
                }
            }
            dataRows++;
        }

        String summary = String.format("Selected %d columns → sheet '%s' (%d rows)",
                step.getInputColumns().size(), outSheetName, dataRows);
        log.debug("SELECT_COLUMNS step {}: {}", step.getOrder(), summary);

        context.setActiveSheet(outSheetName);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .columns(step.getInputColumns())
                .rowCount(dataRows)
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }

    private void copyCellValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case STRING -> target.setCellValue(source.getStringCellValue());
            case NUMERIC -> target.setCellValue(source.getNumericCellValue());
            case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
            case BLANK -> target.setBlank();
            case FORMULA -> {
                try {
                    target.setCellFormula(source.getCellFormula());
                } catch (Exception ignored) {
                    target.setBlank();
                }
            }
            default -> { /* unsupported cell type, leave blank */ }
        }
    }
}
