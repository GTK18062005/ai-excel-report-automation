package com.aiexcel.platform.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

/**
 * Exception thrown when AI-powered analysis recommendation generation fails.
 * Does not expose internals or stack traces to the client.
 */
@ResponseStatus(HttpStatus.UNPROCESSABLE_ENTITY)
public class AnalysisRecommendationException extends RuntimeException {

    public AnalysisRecommendationException(String message) {
        super(message);
    }

    public AnalysisRecommendationException(String message, Throwable cause) {
        super(message, cause);
    }
}
