package com.github.kumasuke120.excel.util;

import org.jetbrains.annotations.ApiStatus;

/**
 * An exception that indicates an error occurred during object creation
 */
@ApiStatus.Internal
public class ObjectCreationException extends IllegalArgumentException {

    /**
     * Constructs a new {@code ObjectCreationException} with the specified detail message
     *
     * @param message the detail message
     * @param cause   the cause
     */
    public ObjectCreationException(String message, Throwable cause) {
        super(message, cause);
    }

}
