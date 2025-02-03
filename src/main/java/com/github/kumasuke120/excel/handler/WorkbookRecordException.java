package com.github.kumasuke120.excel.handler;

/**
 * An exception that indicates an error occurred while processing a workbook record.
 */
public class WorkbookRecordException extends RuntimeException {

    /**
     * Constructs a new {@code WorkbookRecordException} with the specified detail message.
     *
     * @param message the detail message
     */
    WorkbookRecordException(String message) {
        super(message);
    }

    /**
     * Constructs a new {@code WorkbookRecordException} with the specified detail message and cause.
     *
     * @param message the detail message
     * @param cause   the cause
     */
    WorkbookRecordException(String message, Throwable cause) {
        super(message, cause);
    }

}
