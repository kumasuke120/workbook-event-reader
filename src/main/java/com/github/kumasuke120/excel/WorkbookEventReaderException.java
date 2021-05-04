package com.github.kumasuke120.excel;

import org.jetbrains.annotations.NotNull;

/**
 * An abstract {@link RuntimeException} that covers most kinds of exceptions which {@link WorkbookEventReader} may
 * throw
 */
public abstract class WorkbookEventReaderException extends RuntimeException {

    WorkbookEventReaderException(@NotNull Throwable cause) {
        super(cause);
    }

    WorkbookEventReaderException(@NotNull String message) {
        super(message);
    }

    WorkbookEventReaderException(@NotNull String message, Throwable cause) {
        super(message, cause);
    }

}
