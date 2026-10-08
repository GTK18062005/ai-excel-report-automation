package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link CreateChartHandler}.
 * Tests chart creation, validation, error cases, and chained execution compatibility.
 */
class CreateChartHandlerTest {

    @TempDir
    Path tempDir;

    private XSSFWorkbook outputWorkbook;
    private ExecutionContext context;

    @BeforeEach
    void setUp() {
        outputWorkbook = new XSSFWorkbook();
        context = ExecutionContext.builder()
                .outputWorkbook(outputWorkbook)
                .activeSheet(null)
                .steps(List.of())
                .build();
    }

    // -------------------------------------------------------------------------
    // Test data helpers
    // -------------------------------------------------------------------------

    private Path createWorkbookWith(String sheetName, List<String> headers, List<List<Object>> rows) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet(sheetName);
        Row header = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Object val = headers.get(i);
            Cell cell = header.createCell(i);
            if (val instanceof String s) cell.setCellValue(s);
            else if (val instanceof Number n) cell.setCellValue(n.doubleValue());
        }
        for (int ri = 0; ri < rows.size(); ri++) {
            Row row = sheet.createRow(ri + 1);
            List<Object> rowData = rows.get(ri);
            for (int ci = 0; ci < rowData.size(); ci++) {
                Object val = rowData.get(ci);
                Cell cell = row.createCell(ci);
                if (val == null) {
                    // leave as blank
                } else if (val instanceof String s) {
                    cell.setCellValue(s);
                } else if (val instanceof Number n) {
                    cell.setCellValue(n.doubleValue());
                }
            }
        }
        Path path = tempDir.resolve(sheetName + "_input.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    private ExecutionContext buildContext(Path inputPath, String activeSheetName) throws Exception {
        Workbook input = WorkbookFactory.create(inputPath.toFile(), null, true);
        return ExecutionContext.builder()
                .inputWorkbook(input)
                .outputWorkbook(outputWorkbook)
                .activeSheet(activeSheetName)
                .sourcePath(inputPath.toString())
                .steps(List.of())
                .build();
    }

    // -------------------------------------------------------------------------
    // Happy-path: BAR
    // -------------------------------------------------------------------------

    @Test
    void execute_barChart_succeeds() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200),
                        List.of("East", 150)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "BAR",
                        "labelColumn", "Region",
                        "dataColumn", "Amount"))
                .build();

        StepExecutionResult result = new CreateChartHandler().execute(step, context);

        assertEquals(1, result.getOrder());
        assertEquals("CREATE_CHART", result.getOperation());
        assertEquals("Sales_chart_data", result.getSheetName());
        assertEquals(List.of("Region", "Amount"), result.getColumns());
        assertEquals(3, result.getRowCount());
        assertTrue(result.getSummary().contains("BAR"));
        assertTrue(result.getSummary().contains("3 points"));
        assertEquals("Sales_chart_data", result.getNextActiveSheet());
    }

    // -------------------------------------------------------------------------
    // Happy-path: LINE
    // -------------------------------------------------------------------------

    @Test
    void execute_lineChart_succeeds() throws Exception {
        Path path = createWorkbookWith("Data",
                List.of("Month", "Revenue"),
                List.of(
                        List.of("Jan", 1200),
                        List.of("Feb", 1100),
                        List.of("Mar", 1400)));

        context = buildContext(path, "Data");

        WorkflowStep step = WorkflowStep.builder()
                .order(2)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Data",
                        "chartType", "LINE",
                        "labelColumn", "Month",
                        "dataColumn", "Revenue"))
                .build();

        StepExecutionResult result = new CreateChartHandler().execute(step, context);

        assertEquals(2, result.getOrder());
        assertEquals("CREATE_CHART", result.getOperation());
        assertEquals("Data_chart_data", result.getSheetName());
        assertEquals(3, result.getRowCount());
        assertTrue(result.getSummary().contains("LINE"));
    }

    // -------------------------------------------------------------------------
    // Happy-path: PIE
    // -------------------------------------------------------------------------

    @Test
    void execute_pieChart_succeeds() throws Exception {
        Path path = createWorkbookWith("Budget",
                List.of("Category", "Spent"),
                List.of(
                        List.of("IT", 5000),
                        List.of("HR", 3000),
                        List.of("Marketing", 7000)));

        context = buildContext(path, "Budget");

        WorkflowStep step = WorkflowStep.builder()
                .order(3)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Budget",
                        "chartType", "PIE",
                        "labelColumn", "Category",
                        "dataColumn", "Spent"))
                .build();

        StepExecutionResult result = new CreateChartHandler().execute(step, context);

        assertEquals("CREATE_CHART", result.getOperation());
        assertEquals("Budget_chart_data", result.getSheetName());
        assertEquals(3, result.getRowCount());
        assertTrue(result.getSummary().contains("PIE"));
    }

    // -------------------------------------------------------------------------
    // Happy-path: with optional title
    // -------------------------------------------------------------------------

    @Test
    void execute_barChart_withTitle_succeeds() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "BAR",
                        "labelColumn", "Region",
                        "dataColumn", "Amount",
                        "title", "Regional Sales"))
                .build();

        StepExecutionResult result = new CreateChartHandler().execute(step, context);

        assertTrue(result.getSummary().contains("Regional Sales"));
    }

    // -------------------------------------------------------------------------
    // Error: missing source sheet
    // -------------------------------------------------------------------------

    @Test
    void execute_missingSourceSheet_throws() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "NonExistentSheet",
                        "chartType", "BAR",
                        "labelColumn", "Region",
                        "dataColumn", "Amount"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("not found"));
    }

    // -------------------------------------------------------------------------
    // Error: missing label column
    // -------------------------------------------------------------------------

    @Test
    void execute_missingLabelColumn_throws() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "BAR",
                        "labelColumn", "NonExistentLabel",
                        "dataColumn", "Amount"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("labelColumn"));
        assertTrue(ex.getMessage().contains("not found"));
    }

    // -------------------------------------------------------------------------
    // Error: missing data column
    // -------------------------------------------------------------------------

    @Test
    void execute_missingDataColumn_throws() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "BAR",
                        "labelColumn", "Region",
                        "dataColumn", "NonExistentData"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("dataColumn"));
        assertTrue(ex.getMessage().contains("not found"));
    }

    // -------------------------------------------------------------------------
    // Error: invalid chart type
    // -------------------------------------------------------------------------

    @Test
    void execute_invalidChartType_throws() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "SCATTER",
                        "labelColumn", "Region",
                        "dataColumn", "Amount"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("unsupported chartType"));
        assertTrue(ex.getMessage().contains("SCATTER"));
    }

    // -------------------------------------------------------------------------
    // Error: no usable numeric data
    // -------------------------------------------------------------------------

    @Test
    void execute_noNumericData_throws() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", "N/A"),
                        List.of("South", "N/A")));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "BAR",
                        "labelColumn", "Region",
                        "dataColumn", "Amount"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("no usable numeric values"));
    }

    // -------------------------------------------------------------------------
    // Error: null/blank parameters
    // -------------------------------------------------------------------------

    @Test
    void execute_nullParameters_throws() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(null)
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("parameters must be provided"));
    }

    @Test
    void execute_blankSourceSheet_throws() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "   ",
                        "chartType", "BAR",
                        "labelColumn", "Region",
                        "dataColumn", "Amount"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("sourceSheet"));
    }

    @Test
    void execute_blankChartType_throws() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "  ",
                        "labelColumn", "Region",
                        "dataColumn", "Amount"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("chartType"));
    }

    // -------------------------------------------------------------------------
    // Chained execution: chart on output workbook sheet (sheet produced by prior step)
    // -------------------------------------------------------------------------

    @Test
    void execute_chartOnOutputWorkbookSheet_succeeds() throws Exception {
        // Sheet "Sales_agg" exists only in output workbook (simulates AGGREGATE having run)
        Sheet aggSheet = outputWorkbook.createSheet("Sales_agg");
        Row header = aggSheet.createRow(0);
        header.createCell(0).setCellValue("Region");
        header.createCell(1).setCellValue("Total");
        Row r1 = aggSheet.createRow(1);
        r1.createCell(0).setCellValue("North");
        r1.createCell(1).setCellValue(500.0);
        Row r2 = aggSheet.createRow(2);
        r2.createCell(0).setCellValue("South");
        r2.createCell(1).setCellValue(300.0);

        // Input workbook is empty / irrelevant
        XSSFWorkbook emptyInput = new XSSFWorkbook();
        emptyInput.createSheet("Other");
        Path emptyPath = tempDir.resolve("empty.xlsx");
        try (FileOutputStream fos = new FileOutputStream(emptyPath.toFile())) {
            emptyInput.write(fos);
        }
        emptyInput.close();

        context = ExecutionContext.builder()
                .inputWorkbook(emptyInput)
                .outputWorkbook(outputWorkbook)
                .activeSheet("Sales_agg")
                .sourcePath(emptyPath.toString())
                .steps(List.of())
                .build();

        WorkflowStep step = WorkflowStep.builder()
                .order(3)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales_agg",
                        "chartType", "LINE",
                        "labelColumn", "Region",
                        "dataColumn", "Total"))
                .build();

        StepExecutionResult result = new CreateChartHandler().execute(step, context);

        assertEquals("CREATE_CHART", result.getOperation());
        assertEquals("Sales_agg_chart_data", result.getSheetName());
        assertEquals(2, result.getRowCount());
        assertEquals("Sales_agg_chart_data", result.getNextActiveSheet());
    }

    // -------------------------------------------------------------------------
    // Original workbook remains unchanged
    // -------------------------------------------------------------------------

    @Test
    void execute_originalWorkbookUnchanged() throws Exception {
        Path path = createWorkbookWith("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200)));

        byte[] originalBytes = Files.readAllBytes(path);

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "BAR",
                        "labelColumn", "Region",
                        "dataColumn", "Amount"))
                .build();

        new CreateChartHandler().execute(step, context);

        // Input file must not be modified
        assertArrayEquals(originalBytes, Files.readAllBytes(path));
    }

    // -------------------------------------------------------------------------
    // .xls rejection — HSSF source sheet is not supported for chart creation
    // -------------------------------------------------------------------------

    @Test
    void execute_xlsInput_throws() throws Exception {
        // Build an HSSF (.xls) input workbook
        org.apache.poi.hssf.usermodel.HSSFWorkbook hssfInput = new org.apache.poi.hssf.usermodel.HSSFWorkbook();
        org.apache.poi.ss.usermodel.Sheet hssfSheet = hssfInput.createSheet("Data");
        org.apache.poi.ss.usermodel.Row h = hssfSheet.createRow(0);
        h.createCell(0).setCellValue("Region");
        h.createCell(1).setCellValue("Amount");
        org.apache.poi.ss.usermodel.Row r1 = hssfSheet.createRow(1);
        r1.createCell(0).setCellValue("North");
        r1.createCell(1).setCellValue(100);
        Path xlsPath = tempDir.resolve("data.xls");
        try (FileOutputStream fos = new FileOutputStream(xlsPath.toFile())) {
            hssfInput.write(fos);
        }
        hssfInput.close();

        // Open the .xls file read-only as input. The output workbook is XSSF.
        try (Workbook input = WorkbookFactory.create(xlsPath.toFile(), null, true)) {
            ExecutionContext xlsContext = ExecutionContext.builder()
                    .inputWorkbook(input)
                    .outputWorkbook(outputWorkbook)
                    .activeSheet("Data")
                    .sourcePath(xlsPath.toString())
                    .steps(List.of())
                    .build();

            WorkflowStep step = WorkflowStep.builder()
                    .order(1)
                    .operation("CREATE_CHART")
                    .parameters(Map.of(
                            "sourceSheet", "Data",
                            "chartType", "BAR",
                            "labelColumn", "Region",
                            "dataColumn", "Amount"))
                    .build();

            // The source sheet is HSSF (non-XSSF). The handler should fail
            // because chart creation requires an XSSFSheet.
            WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                    () -> new CreateChartHandler().execute(step, xlsContext));

            // Acceptable: error mentions "not found" (input is read-only so sheet
            // resolution surfaces a different error path) or "xlsx" rejection
            String msg = ex.getMessage().toLowerCase();
            assertTrue(msg.contains("not found") || msg.contains("xlsx"),
                    "Expected .xls rejection message, got: " + ex.getMessage());
        }
    }

    // -------------------------------------------------------------------------
    // getSupportedOperation
    // -------------------------------------------------------------------------

    @Test
    void getSupportedOperation_returnsCreateChart() {
        assertEquals("CREATE_CHART", new CreateChartHandler().getSupportedOperation());
    }

    // -------------------------------------------------------------------------
    // AI-proposal compatibility: inputSheets/inputColumns fallback to parameters
    // -------------------------------------------------------------------------

    @Test
    void execute_aiStylePayload_fallsBackToInputSheetsAndColumns() throws Exception {
        // Simulates what the AI emits: sheet/columns on the step, only chartType + title
        // inside parameters. labelColumn / dataColumn / sourceSheet must be derived.
        Path path = createWorkbookWith("Sales",
                List.of("Product", "Sales Amount"),
                List.of(
                        List.of("Apples", 100),
                        List.of("Oranges", 200),
                        List.of("Pears",  150)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .inputSheets(List.of("Sales"))
                .inputColumns(List.of("Product", "Sales Amount"))
                .parameters(Map.of(
                        "chartType", "Bar",
                        "title", "Sales by Product"))
                .build();

        StepExecutionResult result = new CreateChartHandler().execute(step, context);

        assertEquals("CREATE_CHART", result.getOperation());
        assertEquals("Sales_chart_data", result.getSheetName());
        assertEquals(List.of("Product", "Sales Amount"), result.getColumns());
        assertEquals(3, result.getRowCount());
        // chartType was lower-case "Bar" — handler normalises to BAR
        assertTrue(result.getSummary().contains("BAR"));
        assertTrue(result.getSummary().contains("Sales by Product"));
        assertEquals("Sales_chart_data", result.getNextActiveSheet());
    }

    @Test
    void execute_explicitLabelAndDataParams_takePrecedenceOverInputColumns() throws Exception {
        // Explicit parameters must win over inputColumns fallback, even when
        // the inputColumns list is different. This also covers a case where
        // inputColumns is null/empty but the explicit params carry the values.
        Path path = createWorkbookWith("Sales",
                List.of("Product", "Revenue", "Cost"),
                List.of(
                        List.of("Apples", 100, 30),
                        List.of("Oranges", 200, 50),
                        List.of("Pears",  150, 40)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .inputSheets(List.of("Sales"))
                .inputColumns(List.of("Cost"))   // would map to labelColumn[0] if fallback ran
                .parameters(Map.of(
                        "sourceSheet", "Sales",
                        "chartType", "BAR",
                        "labelColumn", "Product",  // explicit
                        "dataColumn", "Revenue")) // explicit
                .build();

        StepExecutionResult result = new CreateChartHandler().execute(step, context);

        // Explicit labelColumn / dataColumn must be used (Product, Revenue)
        assertEquals(List.of("Product", "Revenue"), result.getColumns());
        assertEquals(3, result.getRowCount());
        assertTrue(result.getSummary().contains("labels from 'Product'"));
        assertTrue(result.getSummary().contains("values from 'Revenue'"));
    }

    @Test
    void execute_aiStylePayload_singleInputColumn_dataColumnStillFallsBackFromParamsOrThrows() throws Exception {
        // Only one input column is provided; there is no second column to fall
        // back to for dataColumn. Without an explicit dataColumn, the handler
        // must raise a clear, validation-style error (preserves existing semantics).
        Path path = createWorkbookWith("Sales",
                List.of("Product", "Amount"),
                List.of(
                        List.of("Apples", 100),
                        List.of("Oranges", 200)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .inputSheets(List.of("Sales"))
                .inputColumns(List.of("Product"))  // only one column
                .parameters(Map.of("chartType", "BAR"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("dataColumn"));
    }

    @Test
    void execute_aiStylePayload_noInputSheetsAtAll_stillRequiresSourceSheet() throws Exception {
        // When neither parameters.sourceSheet nor step.inputSheets provide a
        // sheet name, the handler must throw the same validation error a
        // missing parameter would.
        Path path = createWorkbookWith("Sales",
                List.of("Product", "Amount"),
                List.of(List.of("Apples", 100)));

        context = buildContext(path, "Sales");

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("CREATE_CHART")
                .inputSheets(null)
                .inputColumns(List.of("Product", "Amount"))
                .parameters(Map.of("chartType", "BAR"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new CreateChartHandler().execute(step, context));

        assertTrue(ex.getMessage().contains("sourceSheet"));
    }
}
