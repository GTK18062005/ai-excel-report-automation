package com.aiexcel.platform.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

/**
 * Exception thrown when deterministic workflow validation fails due to an
 * unexpected internal error (e.g., the file cannot be read to resolve sheet
 * and column names).
 * <p>
 * Unlike {@link WorkflowProposalException} which carries AI-generated error
 * messages, this exception signals a backend infrastructure problem and is
 * converted to a 422 response by {@link GlobalExceptionHandler}.
 */
@ResponseStatus(HttpStatus.UNPROCESSABLE_ENTITY)
public class WorkflowValidationException extends RuntimeException {

    public WorkflowValidationException(String message) {
        super(message);
    }

    public WorkflowValidationException(String message, Throwable cause) {
        super(message, cause);
    }
}
