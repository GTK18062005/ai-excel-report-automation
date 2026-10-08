package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * Analysis metadata for a single worksheet in a workbook.
 */
@Data
@Builder
public class SheetAnalysis {

    /** Name of the worksheet */
    private String sheetName;

    /** Total number of rows (including header) */
    private int rowCount;

    /** Number of columns */
    private int columnCount;

    /** Analysis for each column in the worksheet */
    private List<ColumnAnalysis> columns;
}
