package com.aiexcel.platform.service;

import com.aiexcel.platform.dto.ColumnAnalysis;
import com.aiexcel.platform.dto.ColumnDataType;
import com.aiexcel.platform.dto.SheetAnalysis;
import com.aiexcel.platform.dto.WorkbookAnalysisResponse;
import com.aiexcel.platform.dto.WorkbookPreviewResponse;
import com.aiexcel.platform.entity.FileStatus;
import com.aiexcel.platform.entity.UploadedFile;
import com.aiexcel.platform.entity.User;
import com.aiexcel.platform.exception.FileNotFoundException;
import com.aiexcel.platform.exception.WorkbookAnalysisException;
import com.aiexcel.platform.repository.UploadedFileRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Deterministic Excel Workbook Analyzer.
 * Reads uploaded workbooks using Apache POI and returns structured metadata.
 * Does NOT use AI — all analysis is performed by deterministic Java code.
 */
@Service
@RequiredArgsConstructor
@Slf4j
public class ExcelAnalyzerService {

    private static final int MAX_SAMPLE_VALUES = 5;
    private static final int MAX_ROWS_TO_ANALYZE = 50000; // Safety limit

    private final UploadedFileRepository fileRepository;

    /**
     * Analyzes an uploaded Excel workbook.
     * Verifies ownership, opens the file, and returns structured metadata.
     *
     * @param fileId ID of the uploaded file
     * @param user   authenticated user (must own the file)
     * @return structured workbook analysis
     */
    @Transactional(readOnly = true)
    public WorkbookAnalysisResponse analyzeWorkbook(Long fileId, User user) {
        log.info("Starting workbook analysis for fileId={}, userId={}", fileId, user.getId());

        // Verify ownership
        UploadedFile file = fileRepository
                .findByIdAndUserAndStatusNot(fileId, user, FileStatus.DELETED)
                .orElseThrow(() -> new FileNotFoundException("File not found with id: " + fileId));

        // Verify physical file exists
        Path filePath = Paths.get(file.getStoragePath());
        if (!Files.exists(filePath)) {
            log.error("Physical file not found for fileId={}, path={}", fileId, file.getStoragePath());
            throw new WorkbookAnalysisException("Physical file not found for ID: " + fileId);
        }

        // Analyze based on file type
        try {
            String fileType = file.getFileType().toLowerCase();
            WorkbookAnalysisResponse response;

            if ("xlsx".equals(fileType) || "xls".equals(fileType)) {
                response = analyzeExcelFile(file, filePath);
            } else if ("csv".equals(fileType)) {
                // For now, CSV analysis is left for the next task
                throw new WorkbookAnalysisException("CSV analysis is not yet implemented");
            } else {
                throw new WorkbookAnalysisException("Unsupported file type: " + fileType);
            }

            log.info("Completed workbook analysis for fileId={}, sheetCount={}",
                    fileId, response.getSheetCount());
            return response;

        } catch (WorkbookAnalysisException e) {
            log.error("Workbook analysis failed for fileId={}: {}", fileId, e.getMessage());
            throw e;
        } catch (Exception e) {
            log.error("Unexpected error during workbook analysis for fileId={}", fileId, e);
            throw new WorkbookAnalysisException("Failed to analyze workbook", e);
        }
    }

    /**
     * Analyzes an Excel file (.xlsx or .xls) using Apache POI.
     */
    private WorkbookAnalysisResponse analyzeExcelFile(UploadedFile file, Path filePath) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(filePath.toFile(), /* password= */ null, /* readOnly= */ true)) {

            int sheetCount = workbook.getNumberOfSheets();
            List<SheetAnalysis> sheetAnalyses = new ArrayList<>();

            for (int i = 0; i < sheetCount; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                SheetAnalysis sheetAnalysis = analyzeSheet(sheet, i);
                sheetAnalyses.add(sheetAnalysis);
            }

            return WorkbookAnalysisResponse.builder()
                    .fileId(file.getId())
                    .fileName(file.getOriginalFileName())
                    .sheetCount(sheetCount)
                    .sheets(sheetAnalyses)
                    .build();
        }
    }

    /**
     * Analyzes a single worksheet.
     */
    private SheetAnalysis analyzeSheet(Sheet sheet, int sheetIndex) {
        String sheetName = sheet.getSheetName() != null ? sheet.getSheetName() : "Sheet" + (sheetIndex + 1);

        // Determine dimensions
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

        // Handle blank sheet
        if (firstRow < 0 || lastRow < 0) {
            return SheetAnalysis.builder()
                    .sheetName(sheetName)
                    .rowCount(0)
                    .columnCount(0)
                    .columns(new ArrayList<>())
                    .build();
        }

        // Limit rows for safety
        int actualLastRow = Math.min(lastRow, firstRow + MAX_ROWS_TO_ANALYZE - 1);
        int rowCount = actualLastRow - firstRow + 1;

        // Read header row (first row)
        Row headerRow = sheet.getRow(firstRow);
        int columnCount = 0;
        String[] headers = new String[0];

        if (headerRow != null) {
            columnCount = headerRow.getLastCellNum();
            if (columnCount > 0) {
                headers = new String[columnCount];
                for (int col = 0; col < columnCount; col++) {
                    Cell cell = headerRow.getCell(col);
                    headers[col] = getCellValueAsString(cell);
                }
            }
        }

        // Analyze each column
        List<ColumnAnalysis> columnAnalyses = new ArrayList<>();
        for (int col = 0; col < columnCount; col++) {
            String columnName = (headers[col] != null && !headers[col].trim().isEmpty())
                    ? headers[col].trim()
                    : "Column_" + (col + 1);

            ColumnAnalysis columnAnalysis = analyzeColumn(sheet, col, firstRow, actualLastRow, columnName);
            columnAnalyses.add(columnAnalysis);
        }

        return SheetAnalysis.builder()
                .sheetName(sheetName)
                .rowCount(rowCount)
                .columnCount(columnCount)
                .columns(columnAnalyses)
                .build();
    }

    /**
     * Analyzes a single column in a worksheet.
     */
    private ColumnAnalysis analyzeColumn(Sheet sheet, int columnIndex, int firstRow, int lastRow, String columnName) {
        long nonEmptyCount = 0;
        long emptyCount = 0;
        Map<String, Long> valueFrequency = new HashMap<>();
        List<String> sampleValues = new ArrayList<>();
        Set<ColumnDataType> detectedTypes = new HashSet<>();

        for (int rowNum = firstRow + 1; rowNum <= lastRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                emptyCount++;
                continue;
            }

            Cell cell = row.getCell(columnIndex);
            if (isCellEmpty(cell)) {
                emptyCount++;
            } else {
                nonEmptyCount++;
                String value = getCellValueAsString(cell);
                ColumnDataType type = detectCellType(cell);

                detectedTypes.add(type);

                // Track frequency for unique/duplicate calculation
                String normalized = value != null ? value.trim() : "";
                if (!normalized.isEmpty()) {
                    valueFrequency.merge(normalized, 1L, Long::sum);
                }

                // Collect sample values (bounded)
                if (sampleValues.size() < MAX_SAMPLE_VALUES) {
                    sampleValues.add(value);
                }
            }
        }

        // Calculate statistics
        long uniqueCount = valueFrequency.size();
        long duplicateCount = valueFrequency.values().stream()
                .filter(count -> count > 1)
                .mapToLong(count -> count - 1) // Each duplicate value beyond the first counts as a duplicate
                .sum();

        double missingPercentage = 0.0;
        long totalCells = nonEmptyCount + emptyCount;
        if (totalCells > 0) {
            missingPercentage = (emptyCount * 100.0) / totalCells;
            missingPercentage = Math.round(missingPercentage * 100.0) / 100.0; // Round to 2 decimal places
        }

        // Determine column type
        ColumnDataType columnType = determineColumnType(detectedTypes, nonEmptyCount);

        return ColumnAnalysis.builder()
                .columnIndex(columnIndex)
                .columnName(columnName)
                .detectedType(columnType)
                .nonEmptyCount(nonEmptyCount)
                .emptyCount(emptyCount)
                .missingPercentage(missingPercentage)
                .uniqueCount(uniqueCount)
                .duplicateCount(duplicateCount)
                .sampleValues(sampleValues)
                .build();
    }

    /**
     * Determines the final column type based on detected types from all cells.
     * Priority: if all cells are empty → EMPTY
     * If single type → that type
     * If mixed types → UNKNOWN
     */
    private ColumnDataType determineColumnType(Set<ColumnDataType> detectedTypes, long nonEmptyCount) {
        if (nonEmptyCount == 0) {
            return ColumnDataType.EMPTY;
        }

        if (detectedTypes.size() == 1) {
            return detectedTypes.iterator().next();
        }

        // Mixed types → UNKNOWN
        return ColumnDataType.UNKNOWN;
    }

    /**
     * Detects the type of a single cell.
     */
    private ColumnDataType detectCellType(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return ColumnDataType.EMPTY;
        }

        CellType cellType = cell.getCellType();

        if (cellType == CellType.BOOLEAN) {
            return ColumnDataType.BOOLEAN;
        }

        if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return ColumnDataType.DATE;
            }
            double value = cell.getNumericCellValue();
            if (value == Math.floor(value) && !Double.isInfinite(value)) {
                return ColumnDataType.INTEGER;
            }
            return ColumnDataType.DECIMAL;
        }

        if (cellType == CellType.STRING) {
            return ColumnDataType.TEXT;
        }

        return ColumnDataType.UNKNOWN;
    }

    /**
     * Gets a cell value as a string representation.
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        CellType cellType = cell.getCellType();

        if (cellType == CellType.BLANK) {
            return "";
        }

        if (cellType == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        }

        if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue().toString();
            }
            double value = cell.getNumericCellValue();
            // Use CellReference to avoid precision issues
            return String.valueOf(value);
        }

        if (cellType == CellType.STRING) {
            String str = cell.getStringCellValue();
            return str != null ? str : "";
        }

        return "";
    }

    /**
     * Checks if a cell is empty or null.
     */
    private boolean isCellEmpty(Cell cell) {
        if (cell == null) {
            return true;
        }

        if (cell.getCellType() == CellType.BLANK) {
            return true;
        }

        if (cell.getCellType() == CellType.STRING) {
            String value = cell.getStringCellValue();
            return value == null || value.trim().isEmpty();
        }

        return false;
    }

    /**
     * Returns a paginated preview of rows from a specific worksheet.
     * Opens the workbook read-only and reads only the requested page's rows.
     *
     * @param fileId   ID of the uploaded file
     * @param user     authenticated user (must own the file)
     * @param sheetName name of the sheet to preview
     * @param page     0-based page number
     * @param pageSize number of rows per page (must be &le; 100)
     * @return paginated sheet preview
     */
    private static final int MAX_PAGE_SIZE = 100;

    @Transactional(readOnly = true)
    public WorkbookPreviewResponse getSheetPreview(Long fileId, User user, String sheetName, int page, int pageSize) {
        // Clamp pageSize to maximum allowed value
        if (pageSize > MAX_PAGE_SIZE) {
            pageSize = MAX_PAGE_SIZE;
        }
        if (pageSize < 1) {
            pageSize = 20;
        }
        if (page < 0) {
            page = 0;
        }

        log.info("Starting sheet preview for fileId={}, sheetName={}, page={}, pageSize={}",
                fileId, sheetName, page, pageSize);

        // Verify ownership
        UploadedFile file = fileRepository
                .findByIdAndUserAndStatusNot(fileId, user, FileStatus.DELETED)
                .orElseThrow(() -> new FileNotFoundException("File not found with id: " + fileId));

        // Verify physical file exists
        Path filePath = Paths.get(file.getStoragePath());
        if (!Files.exists(filePath)) {
            log.error("Physical file not found for fileId={}, path={}", fileId, file.getStoragePath());
            throw new WorkbookAnalysisException("Physical file not found for ID: " + fileId);
        }

        String fileType = file.getFileType().toLowerCase();
        if (!"xlsx".equals(fileType) && !"xls".equals(fileType)) {
            throw new WorkbookAnalysisException("Preview is only supported for .xlsx and .xls files");
        }

        try {
            return readSheetPreview(file, filePath, sheetName, page, pageSize);
        } catch (WorkbookAnalysisException e) {
            throw e;
        } catch (Exception e) {
            log.error("Unexpected error during sheet preview for fileId={}", fileId, e);
            throw new WorkbookAnalysisException("Failed to preview sheet", e);
        }
    }

    private WorkbookPreviewResponse readSheetPreview(UploadedFile file, Path filePath,
                                                     String sheetName, int page, int pageSize) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(filePath.toFile(), null, true)) {

            // Find the sheet by name
            Sheet sheet = null;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet s = workbook.getSheetAt(i);
                if (sheetName.equals(s.getSheetName())) {
                    sheet = s;
                    break;
                }
            }

            if (sheet == null) {
                throw new WorkbookAnalysisException("Sheet not found: " + sheetName);
            }

            int firstRow = sheet.getFirstRowNum();
            int lastRow = sheet.getLastRowNum();

            // Handle empty/blank sheet
            if (firstRow < 0 || lastRow < 0 || lastRow < firstRow) {
                return WorkbookPreviewResponse.builder()
                        .fileId(file.getId())
                        .sheetName(sheetName)
                        .columns(List.of())
                        .rows(List.of())
                        .rowCount(0)
                        .columnCount(0)
                        .page(page)
                        .pageSize(pageSize)
                        .totalPages(0)
                        .build();
            }

            // Read header row for column names
            Row headerRow = sheet.getRow(firstRow);
            int columnCount = 0;
            List<String> columns = new ArrayList<>();

            if (headerRow != null) {
                columnCount = headerRow.getLastCellNum();
                for (int col = 0; col < columnCount; col++) {
                    Cell cell = headerRow.getCell(col);
                    String headerValue = getCellValueAsString(cell);
                    if (headerValue == null || headerValue.trim().isEmpty()) {
                        headerValue = "";
                    } else {
                        headerValue = headerValue.trim();
                    }
                    columns.add(headerValue);
                }
            }

            // Total rows excluding header
            int totalDataRows = lastRow - firstRow;
            int totalPages = (int) Math.ceil((double) totalDataRows / pageSize);
            if (totalPages == 0) {
                totalPages = 1;
            }

            // Build list of data rows (excluding header)
            List<List<String>> pageRows = new ArrayList<>();

            // Data rows start after the header
            int dataStart = firstRow + 1;
            int dataEnd = lastRow;

            // Calculate which rows fall within the requested page
            int pageStartRow = dataStart + (page * pageSize);
            int pageEndRow = Math.min(pageStartRow + pageSize - 1, dataEnd);

            if (pageStartRow <= dataEnd) {
                for (int rowNum = pageStartRow; rowNum <= pageEndRow; rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    List<String> rowValues = new ArrayList<>();
                    if (row != null) {
                        for (int col = 0; col < columnCount; col++) {
                            Cell cell = row.getCell(col);
                            rowValues.add(getCellValueAsString(cell));
                        }
                    }
                    // If row is null or sparse, pad with empty strings
                    while (rowValues.size() < columnCount) {
                        rowValues.add("");
                    }
                    pageRows.add(rowValues);
                }
            }

            return WorkbookPreviewResponse.builder()
                    .fileId(file.getId())
                    .sheetName(sheetName)
                    .columns(columns)
                    .rows(pageRows)
                    .rowCount(totalDataRows)
                    .columnCount(columnCount)
                    .page(page)
                    .pageSize(pageSize)
                    .totalPages(totalPages)
                    .build();
        }
    }
}
