package com.github.kumasuke120.excel.handler;

public class WorkbookRecordException extends RuntimeException {

    WorkbookRecordException(String message) {
        super(message);
    }

    WorkbookRecordException(String message, Throwable cause) {
        super(message, cause);
    }

}
