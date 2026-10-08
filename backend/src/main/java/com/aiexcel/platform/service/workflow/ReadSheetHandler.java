package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

/**
 * Handler for {@code READ_SHEET}.
 * <p>
 * Validates that the requested sheet exists and is readable, then returns
 * metadata about it. This step does not copy data — it establishes the active
 * sheet for subsequent operations.
 */
@Component
@Slf4j
public class ReadSheetHandler implements WorkflowOperationHandler {

    @Override
    public String getSupportedOperation() {
        return "READ_SHEET";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        Workbook workbook = context.getInputWorkbook();
        List<String> inputSheets = step.getInputSheets();

        if (inputSheets == null || inputSheets.isEmpty()) {
            throw new WorkflowExecutionException(
                    "READ_SHEET step " + step.getOrder() + " requires at least one sheet name in inputSheets");
        }

        String targetSheet = inputSheets.get(0);
        Sheet sheet = workbook.getSheet(targetSheet);

        if (sheet == null) {
            throw new WorkflowExecutionException(
                    "READ_SHEET step " + step.getOrder() + ": sheet '" + targetSheet + "' does not exist in workbook");
        }

        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        int rowCount = Math.max(0, lastRow - firstRow + 1);

        int colCount = 0;
        List<String> columns = new ArrayList<>();
        if (firstRow >= 0 && rowCount > 0) {
            var headerRow = sheet.getRow(firstRow);
            if (headerRow != null) {
                colCount = headerRow.getLastCellNum();
                for (int c = 0; c < colCount; c++) {
                    var cell = headerRow.getCell(c);
                    columns.add(cell != null ? cellToString(cell) : "");
                }
            }
        }

        // Set active sheet for subsequent steps
        context.setActiveSheet(targetSheet);

        String summary = String.format(
                "Loaded sheet '%s' (%d rows, %d columns)",
                targetSheet, Math.max(0, rowCount - 1), colCount);

        log.debug("READ_SHEET step {}: {}", step.getOrder(), summary);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(targetSheet)
                .columns(columns)
                .rowCount(Math.max(0, rowCount - 1)) // exclude header
                .summary(summary)
                .nextActiveSheet(targetSheet)
                .build();
    }

    private String cellToString(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue() != null ? cell.getStringCellValue().trim() : "";
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }
}
