package com.aiexcel.platform.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

/**
 * Exception thrown when a workflow approval/rejection cannot be performed
 * (invalid state transition, invalid proposal, etc.).
 * <p>
 * Mapped to HTTP 409 Conflict by {@link GlobalExceptionHandler} since
 * the request was well-formed but the resource is in a state that
 * prevents the requested action.
 */
@ResponseStatus(HttpStatus.CONFLICT)
public class WorkflowApprovalException extends RuntimeException {

    public WorkflowApprovalException(String message) {
        super(message);
    }

    public WorkflowApprovalException(String message, Throwable cause) {
        super(message, cause);
    }
}
