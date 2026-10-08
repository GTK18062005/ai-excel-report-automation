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
import java.nio.file.Path;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Unit tests for {@link ValidateHandler}.
 * Covers: each checkType (not_null, unique, positive) on success and failure,
 * invalid parameters, source preservation, and original-workbook unchanged.
 */
class ValidateHandlerTest {

    @TempDir
    Path tempDir;

    private XSSFWorkbook outputWorkbook;
    private Workbook inputWorkbook;

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

        // Reload as read-only input workbook
        inputWorkbook = WorkbookFactory.create(path.toFile(), null, true);
        return path;
    }

    private ExecutionContext buildContext(String activeSheet) {
        return ExecutionContext.builder()
                .inputWorkbook(inputWorkbook)
                .outputWorkbook(outputWorkbook)
                .activeSheet(activeSheet)
                .sourcePath(tempDir.toString())
                .steps(List.of())
                .build();
    }

    private WorkflowStep buildStep(int order, String checkType, String targetColumn, List<String> inputSheets) {
        return WorkflowStep.builder()
                .order(order)
                .operation("VALIDATE")
                .inputSheets(inputSheets)
                .inputColumns(List.of())
                .parameters(Map.of(
                        "targetColumn", targetColumn,
                        "checkType", checkType))
                .build();
    }

    // -------------------------------------------------------------------------
    // Happy path — not_null
    // -------------------------------------------------------------------------

    @Test
    void execute_notNull_allValuesPresent_returnsPass() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200),
                        List.of("East", 150)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, "not_null", "Region", List.of("Sales"));

        StepExecutionResult result = new ValidateHandler().execute(step, context);

        assertEquals(1, result.getOrder());
        assertEquals("VALIDATE", result.getOperation());
        assertEquals("Sales", result.getSheetName());
        assertEquals(3, result.getRowCount());
        assertTrue(result.getSummary().contains("PASS"));
        assertTrue(result.getSummary().contains("Region"));
        assertEquals("Sales", context.getActiveSheet());
    }

    @Test
    void execute_notNull_emptyValues_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("", 200),
                        List.of("East", 150)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, "not_null", "Region", List.of("Sales"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("Region"));
        assertTrue(ex.getMessage().contains("FAILED"));
    }

    // -------------------------------------------------------------------------
    // Happy path — unique
    // -------------------------------------------------------------------------

    @Test
    void execute_unique_allDistinct_returnsPass() throws Exception {
        createSourceWorkbook("Users",
                List.of("Id", "Name"),
                List.of(
                        List.of(1, "Alice"),
                        List.of(2, "Bob"),
                        List.of(3, "Carol")));

        ExecutionContext context = buildContext("Users");
        WorkflowStep step = buildStep(1, "unique", "Id", List.of("Users"));

        StepExecutionResult result = new ValidateHandler().execute(step, context);

        assertEquals(3, result.getRowCount());
        assertTrue(result.getSummary().contains("PASS"));
        assertTrue(result.getSummary().contains("3 unique values"));
        assertTrue(result.getSummary().contains("0 duplicates"));
    }

    @Test
    void execute_unique_duplicates_throws() throws Exception {
        createSourceWorkbook("Users",
                List.of("Id", "Name"),
                List.of(
                        List.of(1, "Alice"),
                        List.of(2, "Bob"),
                        List.of(1, "Alice2")));

        ExecutionContext context = buildContext("Users");
        WorkflowStep step = buildStep(1, "unique", "Id", List.of("Users"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("Id"));
        assertTrue(ex.getMessage().contains("duplicate"));
    }

    // -------------------------------------------------------------------------
    // Happy path — positive
    // -------------------------------------------------------------------------

    @Test
    void execute_positive_allPositive_returnsPass() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200),
                        List.of("East", 150)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, "positive", "Amount", List.of("Sales"));

        StepExecutionResult result = new ValidateHandler().execute(step, context);

        assertEquals(3, result.getRowCount());
        assertTrue(result.getSummary().contains("PASS"));
        assertTrue(result.getSummary().contains("3 positive values"));
    }

    @Test
    void execute_positive_containsZero_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 0),
                        List.of("East", 150)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, "positive", "Amount", List.of("Sales"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("Amount"));
        assertTrue(ex.getMessage().contains("non-positive"));
    }

    @Test
    void execute_positive_containsNegative_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", -50)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, "positive", "Amount", List.of("Sales"));

        assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
    }

    // -------------------------------------------------------------------------
    // Invalid / missing parameters
    // -------------------------------------------------------------------------

    @Test
    void execute_missingParameters_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = WorkflowStep.builder()
                .order(1).operation("VALIDATE")
                .inputSheets(List.of("Sales"))
                .inputColumns(List.of())
                .parameters(null)
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("parameters must be provided"));
    }

    @Test
    void execute_missingCheckType_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = WorkflowStep.builder()
                .order(1).operation("VALIDATE")
                .inputSheets(List.of("Sales"))
                .inputColumns(List.of())
                .parameters(Map.of("targetColumn", "Region"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("checkType"));
    }

    @Test
    void execute_missingTargetColumn_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = WorkflowStep.builder()
                .order(1).operation("VALIDATE")
                .inputSheets(List.of("Sales"))
                .inputColumns(List.of())
                .parameters(Map.of("checkType", "not_null"))
                .build();

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("targetColumn"));
    }

    @Test
    void execute_unsupportedCheckType_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, "min_length_5", "Region", List.of("Sales"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("unsupported checkType"));
    }

    @Test
    void execute_columnNotFound_throws() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(List.of("North", 100)));

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, "not_null", "NonExistent", List.of("Sales"));

        WorkflowExecutionException ex = assertThrows(WorkflowExecutionException.class,
                () -> new ValidateHandler().execute(step, context));
        assertTrue(ex.getMessage().contains("not found"));
    }

    // -------------------------------------------------------------------------
    // Non-transformative behavior
    // -------------------------------------------------------------------------

    @Test
    void execute_doesNotCreateSheetInOutput() throws Exception {
        createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200)));

        ExecutionContext context = buildContext("Sales");
        int beforeSheetCount = outputWorkbook.getNumberOfSheets();
        WorkflowStep step = buildStep(1, "not_null", "Region", List.of("Sales"));

        new ValidateHandler().execute(step, context);

        int afterSheetCount = outputWorkbook.getNumberOfSheets();
        assertEquals(beforeSheetCount, afterSheetCount,
                "VALIDATE must not create a new sheet in the output workbook");
    }

    @Test
    void execute_doesNotModifyInputWorkbook() throws Exception {
        Path path = createSourceWorkbook("Sales",
                List.of("Region", "Amount"),
                List.of(
                        List.of("North", 100),
                        List.of("South", 200)));

        // Snapshot the input file's first-row data before validation
        long sizeBefore = java.nio.file.Files.size(path);
        long modifiedBefore = java.nio.file.Files.getLastModifiedTime(path).toMillis();

        ExecutionContext context = buildContext("Sales");
        WorkflowStep step = buildStep(1, "unique", "Region", List.of("Sales"));

        new ValidateHandler().execute(step, context);

        // Reload the input from disk to confirm it is unchanged
        try (Workbook reread = WorkbookFactory.create(path.toFile(), null, true)) {
            Sheet sheet = reread.getSheet("Sales");
            assertNotNull(sheet);
            assertEquals("North", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("South", sheet.getRow(2).getCell(0).getStringCellValue());
        }
        assertEquals(sizeBefore, java.nio.file.Files.size(path));
        assertEquals(modifiedBefore, java.nio.file.Files.getLastModifiedTime(path).toMillis());
    }
}
