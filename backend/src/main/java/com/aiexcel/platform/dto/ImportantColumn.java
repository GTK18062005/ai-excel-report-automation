package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

/**
 * A column identified as important by the AI.
 */
@Data
@Builder
public class ImportantColumn {

    /** Column name from the header */
    private String columnName;

    /** Sheet this column belongs to */
    private String sheetName;

    /** Why this column is considered important */
    private String reason;
}
