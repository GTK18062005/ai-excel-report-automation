package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * AI-generated business understanding of an Excel workbook.
 * Derived from structured metadata (not raw data).
 */
@Data
@Builder
public class WorkbookUnderstandingResponse {

    /** ID of the uploaded file these insights were generated for */
    private Long fileId;

    /** Original filename for the uploaded workbook */
    private String fileName;

    /** High-level summary of the workbook content and purpose */
    private String summary;

    /** Inferred business domain (e.g., Sales, HR, Finance, Inventory) */
    private String businessDomain;

    /** Understanding of each sheet */
    private List<SheetUnderstanding> sheets;

    /** Columns deemed most important or impactful */
    private List<ImportantColumn> importantColumns;

    /** Potential relationships between sheets or columns */
    private List<PossibleRelationship> possibleRelationships;

    /** Observations about data quality issues */
    private List<DataQualityObservation> dataQualityObservations;

    /** Suggested analyses that could be performed on this data */
    private List<SuggestedAnalysis> suggestedAnalyses;
}
