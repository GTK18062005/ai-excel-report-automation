package com.aiexcel.platform.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Carries the result of executing a single workflow step.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class StepExecutionResult {

    /** Order of the executed step (1-based). */
    private int order;

    /** Operation name that was executed. */
    private String operation;

    /** Name of the sheet that was produced or used as the primary result. */
    private String sheetName;

    /**
     * Column names referenced in the result, in the order they appear.
     * May be empty for operations that don't preserve column structure.
     */
    private List<String> columns;

    /** Number of rows in the result (excluding header). */
    private int rowCount;

    /**
     * A short, human-readable summary of the step's effect on the workbook.
     */
    private String summary;

    /**
     * The sheet that subsequent steps should operate on by default.
     * May equal {@link #sheetName}.
     */
    private String nextActiveSheet;
}
