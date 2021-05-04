package com.github.kumasuke120.excel;

import org.jetbrains.annotations.NotNull;

/**
 * An exception denotes the original value in a {@link CellValue} cannot be cast to desired type
 */
public class CellValueCastException extends RuntimeException {

    CellValueCastException() {
    }

    @SuppressWarnings("SameParameterValue")
    CellValueCastException(@NotNull String message) {
        super(message);
    }

    CellValueCastException(@NotNull Throwable cause) {
        super(cause);
    }

}
