package com.aiexcel.platform.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

/**
 * Exception thrown when workflow execution fails.
 * <p>
 * Mapped to HTTP 409 Conflict by {@link GlobalExceptionHandler} since
 * the request was well-formed but the resource could not be executed
 * (e.g., invalid operation parameters, unsupported file type, or
 * internal processing error).
 */
@ResponseStatus(HttpStatus.CONFLICT)
public class WorkflowExecutionException extends RuntimeException {

    public WorkflowExecutionException(String message) {
        super(message);
    }

    public WorkflowExecutionException(String message, Throwable cause) {
        super(message, cause);
    }
}
