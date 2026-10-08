package com.aiexcel.platform.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Top-level response DTO for AI-generated business analysis recommendations.
 * Built from the deterministic workbook analysis plus the existing AI workbook understanding.
 * The AI is permitted to decide WHICH analyses are useful; the backend remains deterministic
 * and never executes any analysis or generates runnable code.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class AnalysisRecommendationResponse {

    /** ID of the uploaded file these recommendations were generated for */
    private Long fileId;

    /** Original filename (for context) */
    private String fileName;

    /** High-level summary of the recommendation set, in 1-2 sentences */
    private String overview;

    /** Ordered list of business analysis recommendations */
    private List<AnalysisRecommendation> recommendations;
}
