package com.aiexcel.platform.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

/**
 * Exception thrown when AI-powered workflow proposal generation fails
 * (AI not configured, AI returned malformed output, or AI returned an
 * operation that is not in the allowed catalog).
 * Does not expose internals or stack traces to the client.
 */
@ResponseStatus(HttpStatus.UNPROCESSABLE_ENTITY)
public class WorkflowProposalException extends RuntimeException {

    public WorkflowProposalException(String message) {
        super(message);
    }

    public WorkflowProposalException(String message, Throwable cause) {
        super(message, cause);
    }
}
