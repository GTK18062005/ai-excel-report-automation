package com.aiexcel.platform.dto;

import com.aiexcel.platform.entity.WorkflowStatus;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Response returned after an approve or reject action.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class WorkflowApprovalResponse {

    /** The current status of the workflow. */
    private WorkflowStatus status;

    /**
     * Human-readable summary of what happened.
     * Suitable for direct display to the end user.
     */
    private String message;

    /**
     * The workflow name from the proposal, echoed back for traceability.
     */
    private String workflowName;
}
