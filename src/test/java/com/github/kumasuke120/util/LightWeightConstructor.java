package com.github.kumasuke120.util;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

public class LightWeightConstructor<T> {

    private final Constructor<T> constructor;

    public LightWeightConstructor(Class<T> clazz, Class<?>... parameterTypes) {
        try {
            this.constructor = clazz.getConstructor(parameterTypes);
        } catch (Exception e) {
            throw new AssertionError(e);
        }
    }

    public T newInstance(Object... initArgs) {
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
