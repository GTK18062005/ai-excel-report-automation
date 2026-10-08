package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Handler for {@code CREATE_CHART}.
 * <p>
 * Creates a deterministic chart (BAR, LINE, or PIE) from a source sheet in the
 * output workbook, using Apache POI XDDF APIs already present in {@code poi-ooxml 5.2.5}.
 * The chart is anchored to a sheet in the <b>output</b> workbook.
 * <p>
 * The handler always copies the relevant label and data columns into a new sheet
 * in the output workbook (named {@code <sourceSheet>_chart_data}). This guarantees
 * the chart has a writable target even when the source sheet comes from the
 * read-only input workbook, and ensures the original input file is never modified.
 * <p>
 * Supported parameters in {@code step.parameters}:
 * <ul>
 *   <li>{@code sourceSheet} — required: name of the source sheet</li>
 *   <li>{@code chartType}   — required: one of BAR, LINE, PIE</li>
 *   <li>{@code labelColumn} — required: header name of the column providing category labels</li>
 *   <li>{@code dataColumn}  — required: header name of the column providing numeric values</li>
 *   <li>{@code title}       — optional: chart title</li>
 * </ul>
 */
@Component
@Slf4j
public class CreateChartHandler implements WorkflowOperationHandler {

    private static final Set<String> ALLOWED_CHART_TYPES = Set.of("BAR", "LINE", "PIE");

    @Override
    public String getSupportedOperation() {
        return "CREATE_CHART";
    }

    @Override
    public StepExecutionResult execute(WorkflowStep step, ExecutionContext context) {
        XSSFWorkbook output = (XSSFWorkbook) context.getOutputWorkbook();
        if (output == null) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder()
                            + ": chart creation is only supported for .xlsx output workbooks");
        }

        Map<String, String> params = step.getParameters();
        if (params == null) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": parameters must be provided");
        }

        String chartType = params.get("chartType");
        if (chartType == null || chartType.isBlank()) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": parameter 'chartType' is required");
        }
        chartType = chartType.trim().toUpperCase();
        if (!ALLOWED_CHART_TYPES.contains(chartType)) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": unsupported chartType '"
                            + chartType + "'. Supported: " + String.join(", ", ALLOWED_CHART_TYPES));
        }

        // Resolve sourceSheet: prefer explicit parameter, else fall back to first inputSheets entry.
        String sourceSheet = params.get("sourceSheet");
        if ((sourceSheet == null || sourceSheet.isBlank()) && step.getInputSheets() != null) {
            sourceSheet = step.getInputSheets().stream()
                    .filter(s -> s != null && !s.isBlank())
                    .findFirst()
                    .orElse(null);
        }
        if (sourceSheet == null || sourceSheet.isBlank()) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": parameter 'sourceSheet' is required");
        }

        // Resolve labelColumn: prefer explicit parameter, else fall back to first inputColumns entry.
        String labelColumn = params.get("labelColumn");
        if ((labelColumn == null || labelColumn.isBlank()) && step.getInputColumns() != null
                && !step.getInputColumns().isEmpty()) {
            labelColumn = step.getInputColumns().get(0);
        }
        if (labelColumn == null || labelColumn.isBlank()) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": parameter 'labelColumn' is required");
        }

        // Resolve dataColumn: prefer explicit parameter, else fall back to second inputColumns entry
        // (only if at least two input columns are present).
        String dataColumn = params.get("dataColumn");
        if ((dataColumn == null || dataColumn.isBlank()) && step.getInputColumns() != null
                && step.getInputColumns().size() >= 2) {
            dataColumn = step.getInputColumns().get(1);
        }
        if (dataColumn == null || dataColumn.isBlank()) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": parameter 'dataColumn' is required");
        }

        // Resolve source sheet via the standard utility (input first, then output)
        Sheet resolved = ExecutionUtils.findSheet(context, sourceSheet, step.getOrder());
        if (resolved == null) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder()
                            + ": sheet '" + sourceSheet + "' not found in input or output workbook");
        }
        if (!(resolved instanceof XSSFSheet)) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder()
                            + ": chart creation requires an .xlsx source sheet; '" + sourceSheet
                            + "' is in a legacy .xls workbook");
        }

        // Read header row
        int firstRow = resolved.getFirstRowNum();
        if (firstRow < 0) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder()
                            + ": source sheet '" + sourceSheet + "' is empty");
        }
        Row headerRow = resolved.getRow(firstRow);
        if (headerRow == null) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder()
                            + ": source sheet '" + sourceSheet + "' has no header row");
        }
        int colCount = headerRow.getLastCellNum();
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < colCount; c++) {
            Cell hc = headerRow.getCell(c);
            headers.add(hc != null && hc.getCellType() == CellType.STRING
                    ? hc.getStringCellValue().trim() : "");
        }

        int labelIdx = headers.indexOf(labelColumn);
        if (labelIdx < 0) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": labelColumn '"
                            + labelColumn + "' not found in sheet '" + sourceSheet + "'");
        }

        int dataIdx = headers.indexOf(dataColumn);
        if (dataIdx < 0) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": dataColumn '"
                            + dataColumn + "' not found in sheet '" + sourceSheet + "'");
        }

        int lastRow = resolved.getLastRowNum();
        int numericCount = 0;
        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row row = resolved.getRow(r);
            if (row == null) continue;
            Cell c = row.getCell(dataIdx);
            if (c != null && c.getCellType() == CellType.NUMERIC) {
                numericCount++;
            }
        }
        if (numericCount == 0) {
            throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": no usable numeric values found in column '"
                            + dataColumn + "' of sheet '" + sourceSheet + "'");
        }

        // Create a new sheet in the output workbook to hold the chart + minimal data
        String baseName = sourceSheet + "_chart_data";
        String outSheetName = baseName;
        int suffix = 1;
        while (output.getSheet(outSheetName) != null) {
            outSheetName = baseName + "_" + (++suffix);
        }
        XSSFSheet targetSheet = output.createSheet(outSheetName);

        // header
        Row tHeader = targetSheet.createRow(0);
        tHeader.createCell(0).setCellValue(labelColumn);
        tHeader.createCell(1).setCellValue(dataColumn);
        int outRowIdx = 1;
        for (int r = firstRow + 1; r <= lastRow; r++) {
            Row srcRow = resolved.getRow(r);
            if (srcRow == null) continue;
            Cell labelCell = srcRow.getCell(labelIdx);
            Cell dataCell  = srcRow.getCell(dataIdx);
            if (dataCell == null || dataCell.getCellType() != CellType.NUMERIC) continue;
            Row outRow = targetSheet.createRow(outRowIdx++);
            Cell lbl = outRow.createCell(0);
            if (labelCell != null) {
                switch (labelCell.getCellType()) {
                    case STRING -> lbl.setCellValue(labelCell.getStringCellValue());
                    case NUMERIC -> lbl.setCellValue(labelCell.getNumericCellValue());
                    case BOOLEAN -> lbl.setCellValue(labelCell.getBooleanCellValue());
                    default -> lbl.setCellType(CellType.BLANK);
                }
            }
            Cell val = outRow.createCell(1);
            val.setCellValue(dataCell.getNumericCellValue());
        }
        int dataStartRow = 1;
        int dataEndRow = outRowIdx - 1;

        // Build XDDF data sources anchored to the new target sheet
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                targetSheet, new CellRangeAddress(dataStartRow, dataEndRow, 0, 0));
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                targetSheet, new CellRangeAddress(dataStartRow, dataEndRow, 1, 1));

        // Create drawing + chart anchored to the right of the data
        XSSFDrawing drawing = targetSheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 3, 0, 13, 20);
        XSSFChart chart = drawing.createChart(anchor);

        String title = params.get("title");
        if (title != null && !title.isBlank()) {
            chart.setTitleText(title);
            chart.setTitleOverlay(false);
        }

        switch (chartType) {
            case "BAR" -> {
                XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
                XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
                XDDFBarChartData barData = (XDDFBarChartData) chart.createData(
                        ChartTypes.BAR, bottomAxis, leftAxis);
                XDDFBarChartData.Series series = (XDDFBarChartData.Series) barData.addSeries(categories, values);
                series.setTitle(dataColumn, null);
                chart.plot(barData);
            }
            case "LINE" -> {
                XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
                XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
                XDDFLineChartData lineData = (XDDFLineChartData) chart.createData(
                        ChartTypes.LINE, bottomAxis, leftAxis);
                XDDFLineChartData.Series series = (XDDFLineChartData.Series) lineData.addSeries(categories, values);
                series.setTitle(dataColumn, null);
                series.setSmooth(false);
                chart.plot(lineData);
            }
            case "PIE" -> {
                XDDFChartData pieData = chart.createData(ChartTypes.PIE, null, null);
                pieData.setVaryColors(true);
                XDDFChartData.Series series = pieData.addSeries(categories, values);
                series.setTitle(dataColumn, null);
                chart.plot(pieData);
            }
            default -> throw new WorkflowExecutionException(
                    "CREATE_CHART step " + step.getOrder() + ": unhandled chartType '" + chartType + "'");
        }

        String summary = String.format(
                "CREATE_CHART (%s) on sheet '%s' — labels from '%s', values from '%s' (%d points)%s",
                chartType, outSheetName, labelColumn, dataColumn, numericCount,
                (title != null && !title.isBlank() ? " — title: " + title : ""));
        log.debug("CREATE_CHART step {}: {}", step.getOrder(), summary);

        return StepExecutionResult.builder()
                .order(step.getOrder())
                .operation(step.getOperation())
                .sheetName(outSheetName)
                .columns(List.of(labelColumn, dataColumn))
                .rowCount(numericCount)
                .summary(summary)
                .nextActiveSheet(outSheetName)
                .build();
    }
}
