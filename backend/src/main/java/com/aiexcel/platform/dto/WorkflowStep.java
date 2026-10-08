package com.aiexcel.platform.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * A single step within a proposed workflow.
 * Each step maps to one allowed WorkflowOperation.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WorkflowStep {

    /** 1-based execution order */
    private int order;

    /**
     * Operation name. MUST be one of the values in {@link WorkflowOperation}.
     * The service validates this against the catalog.
     */
    private String operation;

    /** Human-readable description of what this step does */
    private String description;

    /** Sheets this step reads as input */
    private List<String> inputSheets;

    /** Columns this step uses as input */
    private List<String> inputColumns;

    /**
     * Additional parameters for the operation (e.g. filter column, sort direction,
     * aggregation function, chart type). Null is acceptable when no params are needed.
     */
    private Map<String, String> parameters;
}
