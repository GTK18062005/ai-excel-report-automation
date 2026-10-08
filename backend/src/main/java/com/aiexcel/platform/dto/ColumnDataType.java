package com.aiexcel.platform.dto;

/**
 * Detected data type for a column, determined by analyzing actual cell values.
 * Uses deterministic rules — mixed-type columns are classified by majority/compositional logic.
 */
public enum ColumnDataType {
    /** Empty column (no non-empty values) */
    EMPTY,
    /** Whole numeric values (no decimal point) */
    INTEGER,
    /** Decimal numeric values */
    DECIMAL,
    /** Valid date or datetime cells */
    DATE,
    /** Boolean values */
    BOOLEAN,
    /** String/text values */
    TEXT,
    /** Unreadable or ambiguous values that don't fit other types */
    UNKNOWN
}
