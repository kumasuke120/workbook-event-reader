package com.github.kumasuke120.excel.util;

import org.jetbrains.annotations.ApiStatus;

@ApiStatus.Internal
public class ObjectCreationException extends IllegalArgumentException {

    public ObjectCreationException(String message) {
        super(message);
    }

    public ObjectCreationException(String message, Throwable cause) {
        super(message, cause);
    }

}
