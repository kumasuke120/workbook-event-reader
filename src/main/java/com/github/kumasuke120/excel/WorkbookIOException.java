package com.github.kumasuke120.excel;

import org.jetbrains.annotations.NotNull;

/**
 * An exception which related to IO when a {@link WorkbookEventReader} reading a workbook
 */
public class WorkbookIOException extends WorkbookEventReaderException {

    WorkbookIOException(@NotNull String message, @NotNull Throwable cause) {
        super(message, cause);
    }

}
