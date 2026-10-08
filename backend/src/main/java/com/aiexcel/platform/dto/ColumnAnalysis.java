package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * Analysis metadata for a single column in a worksheet.
 * Contains type detection, counts, and a bounded sample of values.
 */
@Data
@Builder
public class ColumnAnalysis {

    /** Zero-based column index */
    private int columnIndex;

    /** Column name from header row (or generated if header is missing) */
    private String columnName;

    /** Detected data type for the column */
    private ColumnDataType detectedType;

    /** Number of non-empty cells in this column */
    private long nonEmptyCount;

    /** Number of empty/blank cells in this column */
    private long emptyCount;

    /** Percentage of empty cells (0-100), rounded to 2 decimal places */
    private double missingPercentage;

    /** Number of unique values in this column */
    private long uniqueCount;

    /** Number of duplicate values in this column (values that appear more than once) */
    private long duplicateCount;

    /** Bounded sample of representative values (max 5) */
    private List<String> sampleValues;
}
