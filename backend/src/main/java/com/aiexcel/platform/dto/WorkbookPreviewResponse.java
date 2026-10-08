package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * Response DTO for paginated sheet preview.
 * Returns a single page of rows from one worksheet.
 */
@Data
@Builder
public class WorkbookPreviewResponse {

    /** ID of the uploaded file */
    private Long fileId;

    /** Name of the worksheet being previewed */
    private String sheetName;

    /** Header column names (from the first row) */
    private List<String> columns;

    /** Rows in the requested page (each row is a list of cell values as strings) */
    private List<List<String>> rows;

    /** Total number of rows in the sheet (including the header row) */
    private int rowCount;

    /** Number of columns in the sheet */
    private int columnCount;

    /** Current page number (0-based) */
    private int page;

    /** Number of rows per page */
    private int pageSize;

    /** Total number of pages */
    private int totalPages;
}
