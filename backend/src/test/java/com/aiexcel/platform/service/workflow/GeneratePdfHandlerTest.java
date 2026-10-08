package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;
import com.aiexcel.platform.exception.WorkflowExecutionException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link GeneratePdfHandler}.
 * Covers: successful PDF generation, missing parameters, empty/no-data,
 * original-workbook unchanged, valid PDF signature, and pagination.
 */
class GeneratePdfHandlerTest {

    @TempDir
    Path tempDir;

    private XSSFWorkbook outputWorkbook;
    private Workbook inputWorkbook;
    private Path sourceXlsxPath;

    @BeforeEach
    void setUp() {
        outputWorkbook = new XSSFWorkbook();
    }

    private Path createSourceWorkbook(String sheetName, List<String> headers, List<List<Object>> rows)
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
                if (val == null) continue;
                if (val instanceof String s) row.createCell(ci).setCellValue(s);
                else if (val instanceof Number n) row.createCell(ci).setCellValue(n.doubleValue());
                else if (val instanceof Boolean b) row.createCell(ci).setCellValue(b);
            }
        }
        Path path = tempDir.resolve(sheetName + "_input.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();

        sourceXlsxPath = path;
        inputWorkbook = WorkbookFactory.create(path.toFile(), null, true);
        return path;
    }

    private ExecutionContext buildContext(String activeSheet) {
        return ExecutionContext.builder()
                .inputWorkbook(inputWorkbook)
                .outputWorkbook(outputWorkbook)
                .activeSheet(activeSheet)
                .sourcePath(sourceXlsxPath.toString())
                .outputPath(sourceXlsxPath.toString()) // output XLSX path; PDF goes beside it
                .steps(List.of())
                .build();
    }

    private WorkflowStep buildStep(int order, Map<String, String> params) {
        return WorkflowStep.builder()
                .order(order)
                .operation("GENERATE_PDF")
                .inputSheets(List.of())
                .inputColumns(List.of())
                .parameters(params)
                .build();
    }

    // -------------------------------------------------------------------------
    // Happy path
    // -------------------------------------------------------------------------

    @Test
    void execute_basicGeneration_writesValidPdf() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200),
                        List.of("East", 150)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, Map.of("title", "Sales Report"));

        StepExecutionResult result = new GeneratePdfHandler().execute(step, context);

        // Result sanity
        assertEquals(1, result.getOrder());
        assertEquals("GENERATE_PDF", result.getOperation());
        assertEquals("Sales", result.getSheetName());
        assertEquals(3, result.getRowCount());
        assertTrue(result.getSummary().contains("GENERATE_PDF"));
        assertTrue(result.getSummary().contains("Sales Report"));

        // PDF file was created at the sibling path
        Path expectedPdf = tempDir.resolve("Sales_input_output.pdf");
        assertTrue(Files.exists(expectedPdf), "PDF should exist at " + expectedPdf);
        assertTrue(Files.size(expectedPdf) > 0, "PDF file should be non-empty");

        // Valid PDF signature
        assertTrue(isValidPdf(expectedPdf), "Generated file should have a valid %PDF- header");

        // Multiple pages wouldn't be expected for 3 rows
        assertTrue(result.getSummary().contains("1 page"),
                "Summary should report 1 page for small dataset, got: " + result.getSummary());
    }

    @Test
    void execute_paginationAcrossMultiplePages() throws Exception {
        // Create 75 rows, default page size 30 → 3 pages
        List<List<Object>> rows = new ArrayList<>();
        for (int i = 0; i < 75; i++) {
            rows.add(List.of("Region" + i, i * 10));
        }
        createSourceWorkbook("Big", List.of("Region", "Amount"), rows);

        ExecutionContext context = buildContext("Big");
        WorkflowStep step = buildStep(1, Map.of());

        StepExecutionResult result = new GeneratePdfHandler().execute(step, context);

        assertEquals(75, result.getRowCount());
        assertTrue(result.getSummary().contains("3 page"),
                "Summary should report 3 pages, got: " + result.getSummary());
    }

    @Test
    void execute_customPageSize_isHonored() throws Exception {
        // 10 rows, pageSize=4 → 3 pages
        List<List<Object>> rows = new ArrayList<>();
        for (int i = 0; i < 10; i++) rows.add(List.of("X" + i, i));
        createSourceWorkbook("Custom", List.of("K", "V"), rows);

        ExecutionContext context = buildContext("Custom");
        WorkflowStep step = buildStep(1, Map.of("pageSize", "4"));

        StepExecutionResult result = new GeneratePdfHandler().execute(step, context);

        assertTrue(result.getSummary().contains("3 page"),
                "Summary should report 3 pages with pageSize=4, got: " + result.getSummary());
    }

    @Test
    void execute_defaultTitle_isApplied() throws Exception {
        createSourceWorkbook("MySheet",
                List.of("A", "B"),
                List.of(List.of(1, 2), List.of(3, 4)));

        ExecutionContext context = buildContext("MySheet");
        WorkflowStep step = buildStep(1, Map.of()); // no title

        StepExecutionResult result = new GeneratePdfHandler().execute(step, context);

        assertTrue(result.getSummary().contains("Workflow Report: MySheet"),
                "Default title should be 'Workflow Report: <sheetName>', got: " + result.getSummary());
    }

    @Test
    void execute_explicitSourceSheet_usedInsteadOfActive() throws Exception {
        // Two sheets: "Main" (active) and "Other" (the one we want)
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet main = wb.createSheet("Main");
        main.createRow(0).createCell(0).setCellValue("MainHeader");
        main.createRow(1).createCell(0).setCellValue("MainData");

        Sheet other = wb.createSheet("Other");
        other.createRow(0).createCell(0).setCellValue("OtherHeader");
        other.createRow(1).createCell(0).setCellValue("OtherData1");
        other.createRow(2).createCell(0).setCellValue("OtherData2");

        Path path = tempDir.resolve("MultiSheet_input.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        sourceXlsxPath = path;
        inputWorkbook = WorkbookFactory.create(path.toFile(), null, true);

        ExecutionContext context = buildContext("Main");
        WorkflowStep step = buildStep(1, Map.of("sourceSheet", "Other"));

        StepExecutionResult result = new GeneratePdfHandler().execute(step, context);

        assertEquals("Other", result.getSheetName());
        assertEquals(2, result.getRowCount());
        assertTrue(result.getSummary().contains("Other"));
    }

    // -------------------------------------------------------------------------
    // Empty / no-data
    // -------------------------------------------------------------------------

    @Test
    void execute_emptySheet_stillProducesPdf() throws Exception {
        // Create a sheet with header but no data rows
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Empty");
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Col1");
        header.createCell(1).setCellValue("Col2");
        Path path = tempDir.resolve("Empty_input.xlsx");
        try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
            wb.write(fos);
        }
        wb.close();
        sourceXlsxPath = path;
        inputWorkbook = WorkbookFactory.create(path.toFile(), null, true);

        ExecutionContext context = buildContext("Empty");
        WorkflowStep step = buildStep(1, Map.of());

        StepExecutionResult result = new GeneratePdfHandler().execute(step, context);

        assertEquals(0, result.getRowCount());
        Path expectedPdf = tempDir.resolve("Empty_input_output.pdf");
        assertTrue(Files.exists(expectedPdf));
        assertTrue(isValidPdf(expectedPdf));
    }

    // -------------------------------------------------------------------------
    // Invalid parameters
    // -------------------------------------------------------------------------

    @Test
    void execute_invalidPageSize_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("A", "B"),
                List.of(List.of(1, 2)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, Map.of("pageSize", "notanumber"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new GeneratePdfHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("pageSize"));
    }

    @Test
    void execute_zeroPageSize_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("A", "B"),
                List.of(List.of(1, 2)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, Map.of("pageSize", "0"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new GeneratePdfHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("pageSize"));
    }

    @Test
    void execute_nullParameters_doesNotThrow() throws Exception {
        // parameters=null is valid (means use all defaults)
        createSourceWorkbook("Sales",
                List.of("A", "B"),
                List.of(List.of(1, 2)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = WorkflowStep.builder()
                .order(1)
                .operation("GENERATE_PDF")
                .inputSheets(List.of())
                .inputColumns(List.of())
                .parameters(null)
                .build();

        StepExecutionResult result = new GeneratePdfHandler().execute(step, context);
        assertEquals(1, result.getRowCount());
    }

    // -------------------------------------------------------------------------
    // Source preservation
    // -------------------------------------------------------------------------

    @Test
    void execute_doesNotModifyInputWorkbook() throws Exception {
        Path path = createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200)));

        long sizeBefore = Files.size(path);
        long modifiedBefore = Files.getLastModifiedTime(path).toMillis();

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, Map.of("title", "Test"));

        new GeneratePdfHandler().execute(step, context);

        // Original file should be unchanged on disk
        assertEquals(sizeBefore, Files.size(path));
        assertEquals(modifiedBefore, Files.getLastModifiedTime(path).toMillis());

        // And the workbook in memory should still have the same content
        try (Workbook reread = WorkbookFactory.create(path.toFile(), null, true)) {
            Sheet sheet = reread.getSheet("Sales");
            assertEquals("North", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals(200.0, sheet.getRow(2).getCell(1).getNumericCellValue());
        }
    }

    @Test
    void execute_doesNotCreateSheetInOutputWorkbook() throws Exception {
        // Output workbook (XLSX) is separate from the PDF file
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200)));

        ExecutionContext context = buildContext("Sales");
        int beforeSheetCount = outputWorkbook.getNumberOfSheets();
        WorkflowStep step = buildStep(1, Map.of());

        new GeneratePdfHandler().execute(step, context);

        // The XLSX output workbook should not have any new sheets added
        int afterSheetCount = outputWorkbook.getNumberOfSheets();
        assertEquals(beforeSheetCount, afterSheetCount,
                "GENERATE_PDF must not add sheets to the output XLSX workbook");
    }

    // -------------------------------------------------------------------------
    // PDF content validity
    // -------------------------------------------------------------------------

    @Test
    void execute_pdfStartsWithPdfMagicBytes() throws Exception {
        createSourceWorkbook("S",
                List.of("A"),
                List.of(List.of("v1"), List.of("v2"), List.of("v3")));

        ExecutionContext context = buildContext("S");
        WorkflowStep step = buildStep(1, Map.of());

        new GeneratePdfHandler().execute(step, context);

        Path pdf = tempDir.resolve("S_input_output.pdf");
        assertTrue(Files.exists(pdf));
        // Read the first 5 bytes; every valid PDF starts with "%PDF-"
        try (InputStream in = Files.newInputStream(pdf)) {
            byte[] head = new byte[5];
            int read = in.read(head);
            assertEquals(5, read, "PDF should be at least 5 bytes");
            assertEquals('%', head[0]);
            assertEquals('P', head[1]);
            assertEquals('D', head[2]);
            assertEquals('F', head[3]);
            assertEquals('-', head[4]);
        }
    }

    @Test
    void execute_pdfFileIsNotEmpty() throws Exception {
        List<List<Object>> rows = new ArrayList<>();
        for (int i = 0; i < 50; i++) rows.add(List.of("row" + i, i));
        createSourceWorkbook("Many", List.of("Label", "Value"), rows);

        ExecutionContext context = buildContext("Many");
        WorkflowStep step = buildStep(1, Map.of());

        new GeneratePdfHandler().execute(step, context);

        Path pdf = tempDir.resolve("Many_input_output.pdf");
        assertTrue(Files.exists(pdf));
        assertTrue(Files.size(pdf) > 100, "Generated PDF should be larger than trivial bytes");
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private boolean isValidPdf(Path path) {
        try (InputStream in = Files.newInputStream(path)) {
            byte[] head = new byte[5];
            int read = in.read(head);
            return read == 5
                    && head[0] == '%' && head[1] == 'P' && head[2] == 'D'
                    && head[3] == 'F' && head[4] == '-';
        } catch (Exception e) {
            return false;
        }
    }
}
