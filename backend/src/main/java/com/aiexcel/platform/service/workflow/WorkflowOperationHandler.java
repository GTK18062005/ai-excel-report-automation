package com.aiexcel.platform.service.workflow;

import com.aiexcel.platform.dto.StepExecutionResult;
import com.aiexcel.platform.dto.WorkflowStep;

/**
 * Strategy interface for executing a single workflow step operation.
 * Each concrete handler handles exactly one {@link com.aiexcel.platform.dto.WorkflowOperation} type.
 */
public interface WorkflowOperationHandler {

    /**
     * Returns the operation name this handler supports (e.g. "READ_SHEET").
     */
    String getSupportedOperation();

    /**
     * Executes the given step against the workbook context.
     * <p>
     * Implementations must NOT close any workbook — the caller owns both.
     * Implementations must NOT modify the input workbook or the original file.
     *
     * @param step    the step to execute
     * @param context the current workbook context (input workbook, current sheet, etc.)
     * @return the result of executing the step; never null
     */
    StepExecutionResult execute(WorkflowStep step, ExecutionContext context);
}
