package com.aiexcel.platform.dto;

import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * Top-level response DTO for workbook analysis.
 * Contains structured metadata about the entire workbook.
 */
@Data
@Builder
public class WorkbookAnalysisResponse {

    /** ID of the uploaded file */
    private Long fileId;

    /** Original filename */
    private String fileName;

    /** Total number of sheets in the workbook */
    private int sheetCount;

    /** Analysis for each sheet in the workbook */
    private List<SheetAnalysis> sheets;
}
