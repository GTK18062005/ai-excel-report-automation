package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

/**
 * AI-generated understanding of a single worksheet.
 */
@Data
@Builder
public class SheetUnderstanding {

    /** Sheet name */
    private String sheetName;

    /** Brief description of what this sheet contains */
    private String description;

    /** Row count as reported by the analyzer */
    private int rowCount;

    /** Column count as reported by the analyzer */
    private int columnCount;
}
