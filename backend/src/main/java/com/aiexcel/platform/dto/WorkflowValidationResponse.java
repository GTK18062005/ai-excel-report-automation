package com.aiexcel.platform.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Deterministic result of validating a proposed workflow.
 * <p>
 * Returned by {@code POST /api/files/{id}/workflow-proposal/validate}.
 * {@code valid} is true only when {@code errors} is empty.
 * Each entry in {@code errors} is a human-readable, action-oriented message
 * that the client can surface directly to the user.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WorkflowValidationResponse {

    /** True when the proposed workflow passes every validation rule. */
    private boolean valid;

    /**
     * Ordered list of validation errors. Empty when the workflow is valid.
     * Messages follow the pattern {@code "step N: <rule> - <reason>"} so the
     * user can locate the offending step at a glance.
     */
    private List<String> errors;
}
