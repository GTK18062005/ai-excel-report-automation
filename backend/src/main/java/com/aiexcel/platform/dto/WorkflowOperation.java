package com.aiexcel.platform.dto;

import java.util.Set;

/**
 * Allowed workflow operations.
 * The AI must ONLY propose operations from this catalog.
 * No custom operations are permitted.
 */
public enum WorkflowOperation {

    READ_SHEET,
    SELECT_COLUMNS,
    REMOVE_EMPTY_ROWS,
    REMOVE_DUPLICATES,
    FILTER,
    SORT,
    CALCULATE,
    GROUP_BY,
    AGGREGATE,
    JOIN_SHEETS,
    PIVOT,
    CREATE_SUMMARY,
    CREATE_CHART,
    VALIDATE,
    GENERATE_PDF,
    EXPORT_EXCEL;

    /** Canonical set for O(1) membership testing */
    private static final Set<String> ALLOWED_NAMES = Set.of(
            "READ_SHEET", "SELECT_COLUMNS", "REMOVE_EMPTY_ROWS", "REMOVE_DUPLICATES",
            "FILTER", "SORT", "CALCULATE", "GROUP_BY", "AGGREGATE",
            "JOIN_SHEETS", "PIVOT", "CREATE_SUMMARY", "CREATE_CHART",
            "VALIDATE", "GENERATE_PDF", "EXPORT_EXCEL"
    );

    /**
     * Returns true if the given string is a valid operation name.
     * Case-sensitive, matches enum constant names exactly.
     */
    public static boolean isValid(String name) {
        return name != null && ALLOWED_NAMES.contains(name.trim());
    }
}
