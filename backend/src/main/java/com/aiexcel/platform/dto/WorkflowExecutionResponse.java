package com.aiexcel.platform.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Response returned after a workflow execution completes (success or failure).
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WorkflowExecutionResponse {

    /**
     * Whether the execution completed successfully.
     */
    private boolean success;

    /**
     * Human-readable summary of the execution result.
     */
    private String message;

    /**
     * Name of the output file created, relative to the storage root.
     * Present only when execution succeeds.
     */
    private String outputFileName;

    /**
     * Human-readable summary of each executed step.
     * Present only when execution succeeds.
     */
    private List<StepResult> steps;

    /**
     * List of error messages for failed steps.
     * Present only when execution partially or fully fails.
     */
    private List<String> errors;

    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class StepResult {
        private int order;
        private String operation;
        private String description;
        private String resultSummary;
    }
}
