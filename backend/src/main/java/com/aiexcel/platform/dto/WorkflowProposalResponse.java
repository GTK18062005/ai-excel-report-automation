package com.aiexcel.platform.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Structured AI response proposing a workflow to fulfil the user's analysis request.
 * The AI proposes WHAT to do; the backend never executes it.
 * All operations are validated against the allowed {@link WorkflowOperation} catalog
 * before the response is returned to the client.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WorkflowProposalResponse {

    /** ID of the uploaded file this proposal is for */
    private Long fileId;

    /** The original user instruction (echoed back) */
    private String instruction;

    /** Short, human-readable name for the proposed workflow */
    private String workflowName;

    /** Plain-language description of what the workflow accomplishes */
    private String description;

    /** Ordered list of steps that make up the workflow */
    private List<WorkflowStep> steps;
}
