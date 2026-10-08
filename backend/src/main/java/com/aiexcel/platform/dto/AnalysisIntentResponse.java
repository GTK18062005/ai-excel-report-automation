package com.aiexcel.platform.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Structured response from the AI interpreting a natural-language analysis request.
 * The AI decides what analyses are implied; the backend stays deterministic
 * and never executes anything.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class AnalysisIntentResponse {

    /** ID of the uploaded file this intent was parsed for */
    private Long fileId;

    /** The original user instruction (echoed back) */
    private String instruction;

    /**
     * High-level intent classification derived from the instruction
     * (e.g., "COMPARISON", "TREND_ANALYSIS", "SEGMENTATION", "ANOMALY_DETECTION").
     */
    private String intent;

    /**
     * Concrete analyses implied by the instruction, each with type and parameters.
     */
    private List<RequestedAnalysis> requestedAnalyses;

    /** Sheet names referenced by the instruction or needed by the requested analyses */
    private List<String> relevantSheets;

    /** Column names referenced by the instruction or needed by the requested analyses */
    private List<String> relevantColumns;

    /** Constraints extracted from the instruction (e.g., time range, region filter) */
    private List<String> constraints;

    /**
     * A single requested analysis within an intent response.
     */
    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class RequestedAnalysis {

        /** Human-readable name of this analysis step */
        private String name;

        /** Category of the analysis (e.g., COMPARISON, TREND, DISTRIBUTION, CORRELATION, SEGMENTATION) */
        private String analysisType;

        /** What this step would compute or produce */
        private String description;
    }
}
