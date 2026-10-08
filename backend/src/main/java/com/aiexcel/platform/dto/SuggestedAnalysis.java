package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

/**
 * A suggested analysis that could be performed on the workbook.
 */
@Data
@Builder
public class SuggestedAnalysis {

    /** Short title of the suggested analysis */
    private String title;

    /** Description of what this analysis would reveal */
    private String description;

    /** Which sheets/columns are relevant */
    private String relevantEntities;
}
