package com.aiexcel.platform.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

/**
 * Exception thrown when AI-powered analysis intent parsing fails.
 * Does not expose internals or stack traces to the client.
 */
@ResponseStatus(HttpStatus.UNPROCESSABLE_ENTITY)
public class AnalysisIntentException extends RuntimeException {

    public AnalysisIntentException(String message) {
        super(message);
    }

    public AnalysisIntentException(String message, Throwable cause) {
        super(message, cause);
    }
}
