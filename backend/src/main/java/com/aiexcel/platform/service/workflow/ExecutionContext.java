package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.WorkflowStep;
import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * Mutable per-execution context shared between operation handlers.
 * <p>
 * Carries the open input workbook (read-only), the output workbook (writeable),
 * and the active sheet name that subsequent steps will operate on by default.
 * Handlers may set {@link #activeSheet} to change which sheet subsequent steps use.
 */
@Data
@Builder
public class ExecutionContext {

    /**
     * The input workbook (opened in read-only mode for inspection).
     * Never modified by handlers.
     */
    private Workbook inputWorkbook;

    /**
     * The output workbook (writeable). Handlers write transformed data here.
     * The caller creates it before execution starts and closes it after all
     * handlers have run.
     */
    private Workbook outputWorkbook;

    /**
     * Name of the sheet that subsequent steps should use as the default
     * input sheet when not explicitly specified by the step.
     */
    private String activeSheet;

    /**
     * Source file path (used to derive a sibling output path).
     * Never modified.
     */
    private String sourcePath;

    /**
     * Output file path that handlers may write to.
     * Null until the executor creates the output file.
     */
    private String outputPath;

    /**
     * All steps in the workflow, in execution order.
     * Provided so handlers can reference upstream steps if needed.
     */
    private List<WorkflowStep> steps;
}
