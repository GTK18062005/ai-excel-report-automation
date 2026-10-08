package com.aiexcel.platform.dto;

import jakarta.validation.constraints.NotBlank;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Request body for AI workflow proposal generation.
 * The user sends a free-text instruction describing the analysis they want;
 * the AI proposes a structured workflow that only uses allowed operations.
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class WorkflowProposalRequest {

    @NotBlank(message = "Instruction must not be blank")
    private String instruction;
}
