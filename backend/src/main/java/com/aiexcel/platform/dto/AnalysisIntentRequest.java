package com.aiexcel.platform.dto;

import jakarta.validation.constraints.NotBlank;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Request body for natural-language analysis intent parsing.
 * The user sends a free-text instruction describing the analysis they want.
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class AnalysisIntentRequest {

    @NotBlank(message = "Instruction must not be blank")
    private String instruction;
}
