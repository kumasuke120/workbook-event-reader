package com.github.kumasuke120.excel;

/**
 * An abstract {@link RuntimeException} that covers most kinds of exceptions which {@link WorkbookEventReader} may
 * throw
 */
@SuppressWarnings("WeakerAccess")
public abstract class WorkbookEventReaderException extends RuntimeException {

    WorkbookEventReaderException(Throwable cause) {
        super(cause);
    }

    WorkbookEventReaderException(String message) {
        super(message);
    }

    WorkbookEventReaderException(String message, Throwable cause) {
        super(message, cause);
    }

}
