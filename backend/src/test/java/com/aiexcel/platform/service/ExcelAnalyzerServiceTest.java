package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.ColumnDataType;
import com.aiexcel.platform.dto.WorkbookAnalysisResponse;
import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.io.TempDir;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class ExcelAnalyzerServiceTest {

    @Mock
    private UploadedFileRepository fileRepository;

    @TempDir
    Path tempDir;

    private ExcelAnalyzerService excelAnalyzerService;
    private User testUser;

    @BeforeEach
    void setUp() {
        excelAnalyzerService = new ExcelAnalyzerService(fileRepository);
        testUser = User.builder()
                .id(1L)
                .name("Test User")
                .email("test@example.com")
                .password("password")
                .build();
    }

    // ========================================================================
    // Type Detection Tests
    // ========================================================================

    @Test
    void analyzeWorkbook_withIntegers_detectsIntegerType() throws IOException {
        // Create test workbook with integer column
        Path excelFile = createTestWorkbook("IntegerTest",
                new String[]{"Id", "Value"},
                new Object[][]{
                        {1, 100},
                        {2, 200},
                        {3, 300}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        assertNotNull(response);
        assertEquals(1, response.getSheetCount());
        assertEquals("Id", response.getSheets().get(0).getColumns().get(0).getColumnName());
        assertEquals(ColumnDataType.INTEGER, response.getSheets().get(0).getColumns().get(0).getDetectedType());
        assertEquals(ColumnDataType.INTEGER, response.getSheets().get(0).getColumns().get(1).getDetectedType());
    }

    @Test
    void analyzeWorkbook_withDecimals_detectsDecimalType() throws IOException {
        Path excelFile = createTestWorkbook("DecimalTest",
                new String[]{"Price"},
                new Object[][]{
                        {10.5},
                        {20.75},
                        {30.99}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        assertEquals(ColumnDataType.DECIMAL, response.getSheets().get(0).getColumns().get(0).getDetectedType());
    }

    @Test
    void analyzeWorkbook_withBooleans_detectsBooleanType() throws IOException {
        Path excelFile = createTestWorkbook("BooleanTest",
                new String[]{"Active"},
                new Object[][]{
                        {true},
                        {false},
                        {true}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        assertEquals(ColumnDataType.BOOLEAN, response.getSheets().get(0).getColumns().get(0).getDetectedType());
    }

    @Test
    void analyzeWorkbook_withStrings_detectsTextType() throws IOException {
        Path excelFile = createTestWorkbook("TextTest",
                new String[]{"Name", "City"},
                new Object[][]{
                        {"Alice", "New York"},
                        {"Bob", "Los Angeles"},
                        {"Charlie", "Chicago"}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        assertEquals(ColumnDataType.TEXT, response.getSheets().get(0).getColumns().get(0).getDetectedType());
        assertEquals(ColumnDataType.TEXT, response.getSheets().get(0).getColumns().get(1).getDetectedType());
    }

    @Test
    void analyzeWorkbook_withDates_detectsDateType() throws IOException {
        Path excelFile = createTestWorkbookWithDates("DateTest",
                new String[]{"Date"},
                new java.util.Date[]{
                        new java.util.Date(2024 - 1900, 0, 1),  // Jan 1, 2024
                        new java.util.Date(2024 - 1900, 1, 15), // Feb 15, 2024
                        new java.util.Date(2024 - 1900, 2, 20)  // Mar 20, 2024
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        assertEquals(ColumnDataType.DATE, response.getSheets().get(0).getColumns().get(0).getDetectedType());
    }

    @Test
    void analyzeWorkbook_withMixedTypes_detectsUnknownType() throws IOException {
        Path excelFile = createTestWorkbook("MixedTest",
                new String[]{"MixedCol"},
                new Object[][]{
                        {"Text"},
                        {123},
                        {true}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        assertEquals(ColumnDataType.UNKNOWN, response.getSheets().get(0).getColumns().get(0).getDetectedType());
    }

    // ========================================================================
    // Missing Value / Empty Cell Tests
    // ========================================================================

    @Test
    void analyzeWorkbook_withEmptyCells_calculatesMissingPercentage() throws IOException {
        Path excelFile = createTestWorkbookWithBlanks("BlankTest",
                new String[]{"Data"},
                new Object[]{null, "Value", null, "Value2", null});

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        var column = response.getSheets().get(0).getColumns().get(0);
        assertEquals(2, column.getNonEmptyCount());
        assertEquals(3, column.getEmptyCount());
        assertEquals(60.0, column.getMissingPercentage()); // 3/5 * 100 = 60%
    }

    @Test
    void analyzeWorkbook_withBlankColumn_detectsEmptyType() throws IOException {
        Path excelFile = createBlankColumnWorkbook("EmptyColTest",
                new String[]{"Col1", "Col2", "Col3"});

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        var columns = response.getSheets().get(0).getColumns();
        assertEquals(ColumnDataType.TEXT, columns.get(0).getDetectedType()); // Has values
        assertEquals(ColumnDataType.EMPTY, columns.get(1).getDetectedType()); // All blank
        assertEquals(ColumnDataType.TEXT, columns.get(2).getDetectedType()); // Has values
    }

    @Test
    void analyzeWorkbook_withMissingHeader_generatesColumnName() throws IOException {
        Path excelFile = createTestWorkbook("NoHeaderTest",
                new String[]{null, "Name"},
                new Object[][]{
                        {null, "Test"}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        var columns = response.getSheets().get(0).getColumns();
        assertEquals("Column_1", columns.get(0).getColumnName());
        assertEquals("Name", columns.get(1).getColumnName());
    }

    // ========================================================================
    // Unique / Duplicate Calculation Tests
    // ========================================================================

    @Test
    void analyzeWorkbook_withDuplicates_calculatesCorrectly() throws IOException {
        Path excelFile = createTestWorkbook("DuplicateTest",
                new String[]{"Value"},
                new Object[][]{
                        {"Apple"},
                        {"Banana"},
                        {"Apple"},      // duplicate
                        {"Cherry"},
                        {"Banana"}      // duplicate
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        var column = response.getSheets().get(0).getColumns().get(0);
        assertEquals(5, column.getNonEmptyCount());
        assertEquals(3, column.getUniqueCount()); // Apple, Banana, Cherry
        assertEquals(2, column.getDuplicateCount()); // 2 extra occurrences (1 Apple + 1 Banana)
    }

    @Test
    void analyzeWorkbook_withAllUniqueValues_duplicateCountIsZero() throws IOException {
        Path excelFile = createTestWorkbook("UniqueTest",
                new String[]{"Value"},
                new Object[][]{
                        {"A"},
                        {"B"},
                        {"C"},
                        {"D"}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        var column = response.getSheets().get(0).getColumns().get(0);
        assertEquals(4, column.getUniqueCount());
        assertEquals(0, column.getDuplicateCount());
    }

    // ========================================================================
    // Sample Values Tests
    // ========================================================================

    @Test
    void analyzeWorkbook_sampleValuesAreBounded() throws IOException {
        Path excelFile = createTestWorkbook("SampleTest",
                new String[]{"Value"},
                new Object[][]{
                        {"Val1"}, {"Val2"}, {"Val3"}, {"Val4"}, {"Val5"},
                        {"Val6"}, {"Val7"}, {"Val8"}, {"Val9"}, {"Val10"}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        var column = response.getSheets().get(0).getColumns().get(0);
        assertTrue(column.getSampleValues().size() <= 5,
                "Sample values should be bounded to 5 or less");
    }

    // ========================================================================
    // Row/Column Count Tests
    // ========================================================================

    @Test
    void analyzeWorkbook_countsRowsAndColumnsCorrectly() throws IOException {
        Path excelFile = createTestWorkbook("CountsTest",
                new String[]{"Col1", "Col2", "Col3"},
                new Object[][]{
                        {1, 2, 3},
                        {4, 5, 6},
                        {7, 8, 9}
                });

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        var sheet = response.getSheets().get(0);
        assertEquals(4, sheet.getRowCount()); // 3 data rows + 1 header row
        assertEquals(3, sheet.getColumnCount());
    }

    @Test
    void analyzeWorkbook_withMultipleSheets_returnsAllSheets() throws IOException {
        Path excelFile = createMultiSheetWorkbook();

        UploadedFile file = createTestFile(excelFile, "xlsx");
        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        WorkbookAnalysisResponse response = excelAnalyzerService.analyzeWorkbook(1L, testUser);

        assertEquals(2, response.getSheetCount());
        assertEquals("Sales", response.getSheets().get(0).getSheetName());
        assertEquals("Targets", response.getSheets().get(1).getSheetName());
    }

    // ========================================================================
    // Security / Ownership Tests
    // ========================================================================

    @Test
    void analyzeWorkbook_fileNotOwnedByUser_throwsFileNotFoundException() {
        when(fileRepository.findByIdAndUserAndStatusNot(999L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.empty());

        assertThrows(FileNotFoundException.class, () -> {
            excelAnalyzerService.analyzeWorkbook(999L, testUser);
        });
    }

    @Test
    void analyzeWorkbook_csvFile_throwsWorkbookAnalysisException() throws IOException {
        Path csvFile = tempDir.resolve("test.csv");
        java.nio.file.Files.writeString(csvFile, "col1,col2\n1,2\n");

        UploadedFile file = UploadedFile.builder()
                .id(1L)
                .user(testUser)
                .originalFileName("test.csv")
                .storedFileName("test.csv")
                .fileType("csv")
                .storagePath(csvFile.toString())
                .status(FileStatus.UPLOADED)
                .build();

        when(fileRepository.findByIdAndUserAndStatusNot(1L, testUser, FileStatus.DELETED))
                .thenReturn(Optional.of(file));

        assertThrows(Exception.class, () -> {
            excelAnalyzerService.analyzeWorkbook(1L, testUser);
        });
    }

    // ========================================================================
    // Helper Methods
    // ========================================================================

    private Path createTestWorkbook(String sheetName, String[] headers, Object[][] data) throws IOException {
        Path file = tempDir.resolve(sheetName + ".xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);
            int rowNum = 0;

            // Create header row
            Row headerRow = sheet.createRow(rowNum++);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                if (headers[i] != null) {
                    cell.setCellValue(headers[i]);
                }
            }

            // Create data rows
            for (Object[] row : data) {
                Row dataRow = sheet.createRow(rowNum++);
                for (int i = 0; i < row.length; i++) {
                    Cell cell = dataRow.createCell(i);
                    if (row[i] != null) {
                        setCellValue(cell, row[i]);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }
        return file;
    }

    private Path createTestWorkbookWithDates(String sheetName, String[] headers, java.util.Date[] dates) throws IOException {
        Path file = tempDir.resolve(sheetName + ".xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);
            CreationHelper createHelper = workbook.getCreationHelper();
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd"));

            int rowNum = 0;

            // Create header row
            Row headerRow = sheet.createRow(rowNum++);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Create data rows with dates
            for (java.util.Date date : dates) {
                Row dataRow = sheet.createRow(rowNum++);
                Cell cell = dataRow.createCell(0);
                cell.setCellValue(date);
                cell.setCellStyle(dateStyle);
            }

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }
        return file;
    }

    private Path createTestWorkbookWithBlanks(String sheetName, String[] headers, Object[] data) throws IOException {
        Path file = tempDir.resolve(sheetName + ".xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);
            int rowNum = 0;

            // Create header row
            Row headerRow = sheet.createRow(rowNum++);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                if (headers[i] != null) {
                    cell.setCellValue(headers[i]);
                }
            }

            // Create data rows with possible null values
            for (Object value : data) {
                Row dataRow = sheet.createRow(rowNum++);
                Cell cell = dataRow.createCell(0);
                if (value != null) {
                    setCellValue(cell, value);
                }
                // Leave blank if null
            }

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }
        return file;
    }

    private Path createBlankColumnWorkbook(String sheetName, String[] headers) throws IOException {
        Path file = tempDir.resolve(sheetName + ".xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);
            int rowNum = 0;

            // Create header row
            Row headerRow = sheet.createRow(rowNum++);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Create 3 data rows - only first and third column have values
            for (int i = 0; i < 3; i++) {
                Row dataRow = sheet.createRow(rowNum++);
                dataRow.createCell(0).setCellValue("Value");
                // Column 1 is intentionally blank
                dataRow.createCell(2).setCellValue("Text");
            }

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }
        return file;
    }

    private Path createMultiSheetWorkbook() throws IOException {
        Path file = tempDir.resolve("MultiSheet.xlsx");
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create Sales sheet
            Sheet salesSheet = workbook.createSheet("Sales");
            Row salesHeader = salesSheet.createRow(0);
            salesHeader.createCell(0).setCellValue("Product");
            salesHeader.createCell(1).setCellValue("Amount");
            Row salesData = salesSheet.createRow(1);
            salesData.createCell(0).setCellValue("Widget");
            salesData.createCell(1).setCellValue(100);

            // Create Targets sheet
            Sheet targetsSheet = workbook.createSheet("Targets");
            Row targetsHeader = targetsSheet.createRow(0);
            targetsHeader.createCell(0).setCellValue("Product");
            targetsHeader.createCell(1).setCellValue("Target");
            Row targetsData = targetsSheet.createRow(1);
            targetsData.createCell(0).setCellValue("Widget");
            targetsData.createCell(1).setCellValue(150);

            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                workbook.write(fos);
            }
        }
        return file;
    }

    private void setCellValue(Cell cell, Object value) {
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof String[]) {
            cell.setCellValue(String.join(",", (String[]) value));
        }
    }

    private UploadedFile createTestFile(Path filePath, String fileType) {
        return UploadedFile.builder()
                .id(1L)
                .user(testUser)
                .originalFileName(filePath.getFileName().toString())
                .storedFileName(filePath.getFileName().toString())
                .fileType(fileType)
                .fileSize(1024L)
                .storagePath(filePath.toString())
                .status(FileStatus.UPLOADED)
                .uploadedAt(LocalDateTime.now())
                .build();
    }
}
