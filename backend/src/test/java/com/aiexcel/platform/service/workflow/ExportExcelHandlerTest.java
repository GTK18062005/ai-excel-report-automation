package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ExportExcelHandler}.
 * Covers: happy-path export, sheet filtering, custom filename,
 * missing sheet, empty sheets param, original workbook unchanged, and export file validity.
 */
class ExportExcelHandlerTest {

    @TempDir
    Path tempDir;

    private XSSFWorkbook outputWorkbook;

    @BeforeEach
    void setUp() {
        outputWorkbook = new XSSFWorkbook();
    }

    private Path writeInputFile(String sheetName, List<String> headers, List<List<Object>> rows)
            throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet(sheetName);
        Row header = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            header.createCell(i).setCellValue(headers.get(i));
        }
        for (int ri = 0; ri < rows.size(); ri++) {
            Row row = sheet.createRow(ri + 1);
            List<Object> rowData = rows.get(ri);
            for (int ci = 0; ci < rowData.size(); ci++) {
                Object val = rowData.get(ci);
                if (val instanceof String s) row.createCell(ci).setCellValue(s);
                else if (val instanceof Number n) row.createCell(ci).setCellValue(n.doubleValue());
                else if (val instanceof Boolean b) row.createCell(ci).setCellValue(b);
            }
        }
        Path path = tempDir.resolve("input.xlsx");
        try (java.io.FileOutputStream fos = new java.io.FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        return path;
    }

    private ExecutionContext buildContext(Path sourcePath) {
        return ExecutionContext.builder()
                .inputWorkbook(null)
                .outputWorkbook(outputWorkbook)
                .activeSheet(null)
                .sourcePath(sourcePath.toString())
                .outputPath(tempDir.resolve("output.xlsx").toString())
                .steps(List.of())
                .build();
    }

    private WorkflowStep buildStep(int order, Map<String, String> params) {
        return WorkflowStep.builder()
                .order(order)
                .operation("EXPORT_EXCEL")
                .inputSheets(List.of())
                .inputColumns(List.of())
                .parameters(params)
                .build();
    }

    // -------------------------------------------------------------------------
    // Helper: populate the output workbook with named sheets
    // -------------------------------------------------------------------------

    private void addSheetToOutput(String name, List<String> headers, List<List<Object>> rows) {
        Sheet sheet = outputWorkbook.createSheet(name);
        Row header = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            header.createCell(i).setCellValue(headers.get(i));
        }
        for (int ri = 0; ri < rows.size(); ri++) {
            Row row = sheet.createRow(ri + 1);
            List<Object> rowData = rows.get(ri);
            for (int ci = 0; ci < rowData.size(); ci++) {
                Object val = rowData.get(ci);
                if (val instanceof String s) row.createCell(ci).setCellValue(s);
                else if (val instanceof Number n) row.createCell(ci).setCellValue(n.doubleValue());
            }
        }
    }

    // -------------------------------------------------------------------------
    // Happy path — export all sheets
    // -------------------------------------------------------------------------

    @Test
    void execute_exportAllSheets_createsValidXlsx() throws Exception {
        addSheetToOutput("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100), List.of("South", 200)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of());

        StepExecutionResult result = new ExportExcelHandler().execute(step, context);

        assertEquals(1, result.getOrder());
        assertEquals("EXPORT_EXCEL", result.getOperation());
        assertTrue(result.getSummary().contains("EXPORT_EXCEL"));
        assertTrue(result.getSummary().contains("_exported.xlsx"));

        Path exportFile = tempDir.resolve("input_exported.xlsx");
        assertTrue(Files.exists(exportFile), "Export file should exist at " + exportFile);
        assertTrue(Files.size(exportFile) > 0);

        // Verify the export contains the expected sheet and data
        try (Workbook wb = WorkbookFactory.create(exportFile.toFile(), null, true)) {
            assertEquals(1, wb.getNumberOfSheets());
            Sheet sheet = wb.getSheet("Sales");
            assertNotNull(sheet);
            assertEquals("Region", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("North", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(100.0, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.001);
        }
    }

    @Test
    void execute_exportMultipleSheets_allPresentInOutput() throws Exception {
        addSheetToOutput("Main",
                List.of("A", "B"),
                List.of(List.of(1, 2)));
        addSheetToOutput("Aux",
                List.of("X", "Y"),
                List.of(List.of(3, 4)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of());

        StepExecutionResult result = new ExportExcelHandler().execute(step, context);

        Path exportFile = tempDir.resolve("input_exported.xlsx");
        try (Workbook wb = WorkbookFactory.create(exportFile.toFile(), null, true)) {
            assertEquals(2, wb.getNumberOfSheets());
            assertNotNull(wb.getSheet("Main"));
            assertNotNull(wb.getSheet("Aux"));
        }
    }

    // -------------------------------------------------------------------------
    // Sheet filtering
    // -------------------------------------------------------------------------

    @Test
    void execute_filterSheets_onlyRequestedSheetsPresent() throws Exception {
        addSheetToOutput("Keep",
                List.of("Col1"),
                List.of(List.of("v1")));
        addSheetToOutput("Drop",
                List.of("Col2"),
                List.of(List.of("v2")));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of("sheets", "Keep"));

        StepExecutionResult result = new ExportExcelHandler().execute(step, context);

        Path exportFile = tempDir.resolve("input_exported.xlsx");
        try (Workbook wb = WorkbookFactory.create(exportFile.toFile(), null, true)) {
            assertEquals(1, wb.getNumberOfSheets());
            assertNotNull(wb.getSheet("Keep"));
            assertNull(wb.getSheet("Drop"));
        }
    }

    @Test
    void execute_filterMultipleSheets_allPresent() throws Exception {
        addSheetToOutput("S1", List.of("A"), List.of(List.of(1)));
        addSheetToOutput("S2", List.of("B"), List.of(List.of(2)));
        addSheetToOutput("S3", List.of("C"), List.of(List.of(3)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of("sheets", "S1, S3"));

        StepExecutionResult result = new ExportExcelHandler().execute(step, context);

        Path exportFile = tempDir.resolve("input_exported.xlsx");
        try (Workbook wb = WorkbookFactory.create(exportFile.toFile(), null, true)) {
            assertEquals(2, wb.getNumberOfSheets());
            assertNotNull(wb.getSheet("S1"));
            assertNull(wb.getSheet("S2"));
            assertNotNull(wb.getSheet("S3"));
        }
    }

    @Test
    void execute_missingSheet_throws() throws Exception {
        addSheetToOutput("Exists",
                List.of("A"),
                List.of(List.of(1)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of("sheets", "NonExistent"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ExportExcelHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("NonExistent"));
        assertTrue(ex.getMessage().contains("not found"));
    }

    @Test
    void execute_partiallyMissingSheets_throwsWithMissingNames() throws Exception {
        addSheetToOutput("Exists",
                List.of("A"),
                List.of(List.of(1)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of("sheets", "Exists, Missing"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ExportExcelHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("Missing"));
    }

    @Test
    void execute_emptySheetsParam_throws() throws Exception {
        addSheetToOutput("Sales",
                List.of("A"),
                List.of(List.of(1)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of("sheets", "   ,   "));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ExportExcelHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("sheets"));
    }

    // -------------------------------------------------------------------------
    // Custom filename
    // -------------------------------------------------------------------------

    @Test
    void execute_customFilename_isUsed() throws Exception {
        addSheetToOutput("Data",
                List.of("X"),
                List.of(List.of("y")));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of("filename", "my_custom_export.xlsx"));

        StepExecutionResult result = new ExportExcelHandler().execute(step, context);

        Path customFile = tempDir.resolve("my_custom_export.xlsx");
        assertTrue(Files.exists(customFile), "Custom filename should be used");
        assertTrue(result.getSummary().contains("my_custom_export.xlsx"));
    }

    @Test
    void execute_filenameWithoutExtension_getsXlsxAppended() throws Exception {
        addSheetToOutput("Data",
                List.of("X"),
                List.of(List.of("y")));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of("filename", "bare_name"));

        new ExportExcelHandler().execute(step, context);

        Path expected = tempDir.resolve("bare_name.xlsx");
        assertTrue(Files.exists(expected));
    }

    // -------------------------------------------------------------------------
    // Source preservation
    // -------------------------------------------------------------------------

    @Test
    void execute_doesNotModifyOutputWorkbook() throws Exception {
        addSheetToOutput("Original",
                List.of("A", "B"),
                List.of(List.of(1, 2), List.of(3, 4)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        int sheetCountBefore = outputWorkbook.getNumberOfSheets();

        WorkflowStep step = buildStep(1, Map.of());
        new ExportExcelHandler().execute(step, context);

        // Output workbook must be unchanged
        assertEquals(sheetCountBefore, outputWorkbook.getNumberOfSheets());
        assertNotNull(outputWorkbook.getSheet("Original"));
    }

    // -------------------------------------------------------------------------
    // Edge cases
    // -------------------------------------------------------------------------

    @Test
    void execute_nullParameters_doesNotThrow() throws Exception {
        addSheetToOutput("Sheet1",
                List.of("A"),
                List.of(List.of(1)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);

        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("EXPORT_EXCEL")
                .inputSheets(List.of())
                .inputColumns(List.of())
                .parameters(null)
                .build();

        StepExecutionResult result = new ExportExcelHandler().execute(step, context);
        assertEquals("EXPORT_EXCEL", result.getOperation());
    }

    @Test
    void execute_emptyWorkbook_createsValidEmptyXlsx() throws Exception {
        // Empty output workbook (no sheets added)
        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of());

        StepExecutionResult result = new ExportExcelHandler().execute(step, context);

        Path exportFile = tempDir.resolve("input_exported.xlsx");
        assertTrue(Files.exists(exportFile));
        try (Workbook wb = WorkbookFactory.create(exportFile.toFile(), null, true)) {
            assertEquals(0, wb.getNumberOfSheets());
        }
    }

    @Test
    void execute_exportedFileStartsWithPK() throws Exception {
        // XLSX files start with PK (ZIP format)
        addSheetToOutput("Data",
                List.of("A", "B"),
                List.of(List.of(1, 2), List.of(3, 4)));

        Path sourcePath = writeInputFile("dummy", List.of("X"), List.of(List.of("x")));
        ExecutionContext context = buildContext(sourcePath);
        WorkflowStep step = buildStep(1, Map.of());

        new ExportExcelHandler().execute(step, context);

        Path exportFile = tempDir.resolve("input_exported.xlsx");
        byte[] head = Files.readAllBytes(exportFile);
        assertTrue(head.length > 4, "File should have a header");
        assertEquals(0x50, head[0] & 0xFF); // 'P'
        assertEquals(0x4B, head[1] & 0xFF); // 'K'
    }
}
