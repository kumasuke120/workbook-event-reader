package com.github.kumasuke120.excel.util;


import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.lang.reflect.Constructor;

/**
 * A utility class that provides some common operations on object creation
 *
 * @param <T> the type of object to be created
 */
@ApiStatus.Internal
public class ObjectFactory<T> {

    private final Class<T> clazz;
    private final Constructor<T> constructor;
    private final MethodHandle constructorHandle;
    private final boolean useHandle;

    private ObjectFactory(Class<T> clazz,
                          Constructor<T> constructor, MethodHandle constructorHandle,
                          boolean useHandle) {
        this.clazz = clazz;
        this.constructor = constructor;
        this.constructorHandle = constructorHandle;
        this.useHandle = useHandle;
    }

    /**
     * Creates a new {@code ObjectFactory} for the specified class
     *
     * @param clazz the class
     * @param <T>   the type of object to be created
     * @return a new {@code ObjectFactory} for the specified class
     * @throws ObjectCreationException if failed to create the factory, such as the class has no public constructor
     */
    @NotNull
    public static <T> ObjectFactory<T> buildFactory(@NotNull Class<T> clazz) {
        try {
            return buildFactory(clazz, true);
        } catch (ObjectCreationException e1) {
            try {
                return buildFactory(clazz, false);
            } catch (ObjectCreationException e2) {
                e2.addSuppressed(e1);
                throw e2;
            }
        }
    }

    /**
     * Creates a new {@code ObjectFactory} for the specified class
     *
     * @param clazz     the class
     * @param useHandle whether to use {@link MethodHandle} to create the object
     * @param <T>       the type of object to be created
     * @return a new {@code ObjectFactory} for the specified class
     */
    @NotNull
    static <T> ObjectFactory<T> buildFactory(@NotNull Class<T> clazz, boolean useHandle) {
        if (useHandle) {
            return buildFactoryWithHandle(clazz);
        } else {
            return buildFactoryWithReflection(clazz);
        }
    }

    @NotNull
    private static <T> ObjectFactory<T> buildFactoryWithHandle(@NotNull Class<T> clazz) {
        final MethodHandles.Lookup lookup = MethodHandles.lookup();

        final MethodHandle constructorHandle;
        try {
            constructorHandle = lookup.findConstructor(clazz, MethodType.methodType(void.class));
        } catch (NoSuchMethodException | IllegalAccessException e) {
            throw new ObjectCreationException("Failed to find constructor of '" + clazz + "'", e);
        }

        return new ObjectFactory<>(clazz, null, constructorHandle, true);
    }

    @NotNull
    private static <T> ObjectFactory<T> buildFactoryWithReflection(@NotNull Class<T> clazz) {
        final Constructor<T> constructor;
        try {
            constructor = clazz.getConstructor();
        } catch (NoSuchMethodException e) {
            throw new ObjectCreationException("Failed to find constructor of '" + clazz + "'", e);
        }

        return new ObjectFactory<>(clazz, constructor, null, false);
    }

    /**
     * Creates a new instance of the object
     *
     * @return a new instance of the object
     * @throws ObjectCreationException if failed to create the object
     */
    @NotNull
    public T newInstance() {
        if (useHandle) {
            return newInstanceWithHandle();
        } else {
            return newInstanceWithReflection();
        }
    }

    @SuppressWarnings("unchecked")
    private T newInstanceWithHandle() {
        try {
            return (T) constructorHandle.invoke();
        } catch (ClassCastException ce) {
            throw new AssertionError("Shouldn't happen", ce);
        } catch (Throwable e) {
            throw new ObjectCreationException("Failed to create new instance of '" + clazz + "'", e);
        }
    }

    private T newInstanceWithReflection() {
        try {
            return constructor.newInstance();
        } catch (ReflectiveOperationException e) {
            throw new ObjectCreationException("Failed to create new instance of '" + clazz + "'", e);
        }
    }

}
