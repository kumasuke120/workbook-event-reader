package com.github.kumasuke120.util;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

public class LightWeightConstructor<T> {

    private final Constructor<T> constructor;

    public LightWeightConstructor(@NotNull Class<T> clazz, @NotNull Class<?>... parameterTypes) {
        try {
            this.constructor = clazz.getConstructor(parameterTypes);
        } catch (Exception e) {
            throw new AssertionError(e);
        }
    }

    @NotNull
    public T newInstance(@Nullable Object... initArgs) {
        try {
            return constructor.newInstance(initArgs);
        } catch (InvocationTargetException e) {
            final Throwable targetException = e.getTargetException();
            if (targetException instanceof RuntimeException) {
                throw (RuntimeException) targetException;
            } else {
                throw new AssertionError(e);
            }
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

}
