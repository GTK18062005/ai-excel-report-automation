package com.aiexcel.platform.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

/**
 * Exception thrown when workbook analysis fails due to invalid/unreadable file.
 */
@ResponseStatus(HttpStatus.UNPROCESSABLE_ENTITY)
public class WorkbookAnalysisException extends RuntimeException {

    public WorkbookAnalysisException(String message) {
        super(message);
    }

    public WorkbookAnalysisException(String message, Throwable cause) {
        super(message, cause);
    }
}
