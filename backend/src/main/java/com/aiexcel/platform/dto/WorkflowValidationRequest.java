package com.aiexcel.platform.dto;

import jakarta.validation.Valid;
import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Request body for deterministic workflow validation.
 * <p>
 * Mirrors the shape of {@link WorkflowProposalResponse} so that the AI-generated
 * JSON can be forwarded to this endpoint for validation before execution.
 * <p>
 * Only the {@code steps} field is required; {@code instruction}, {@code workflowName},
 * and {@code description} are accepted as-is and echoed in the response for
 * traceability.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WorkflowValidationRequest {

    /** Original user instruction (optional — echoed in response). */
    private String instruction;

    /** Proposed workflow name (optional — echoed in response). */
    private String workflowName;

    /** Plain-language description (optional — echoed in response). */
    private String description;

    /**
     * Ordered list of steps to validate. Null or empty is accepted (returns
     * valid=true with no errors, matching the AI-proposal degenerate case).
     */
    @NotNull(message = "Steps must not be null")
    @Valid
    private List<WorkflowStep> steps;
}
