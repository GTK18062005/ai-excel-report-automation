package com.aiexcel.platform.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

/**
 * Exception thrown when AI-powered workbook understanding fails.
 * Does not expose internals or stack traces to the client.
 */
@ResponseStatus(HttpStatus.UNPROCESSABLE_ENTITY)
public class WorkbookUnderstandingException extends RuntimeException {

    public WorkbookUnderstandingException(String message) {
        super(message);
    }

    public WorkbookUnderstandingException(String message, Throwable cause) {
        super(message, cause);
    }
}
