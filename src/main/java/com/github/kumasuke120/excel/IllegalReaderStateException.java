package com.github.kumasuke120.excel;

import org.jetbrains.annotations.NotNull;

/**
 * An exception occurred when the {@link WorkbookEventReader} stays in a particular state where some kind of
 * operation cannot be performed
 */
public class IllegalReaderStateException extends WorkbookEventReaderException {

    IllegalReaderStateException(@NotNull String message) {
        super(message);
    }

}
