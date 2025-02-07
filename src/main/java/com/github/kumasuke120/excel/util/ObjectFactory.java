package com.github.kumasuke120.excel.util;


import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;

/**
 * A utility class that provides some common operations on object creation
 *
 * @param <T> the type of object to be created
 */
@ApiStatus.Internal
public class ObjectFactory<T> {

    private final Class<T> clazz;
    private final MethodHandle constructorHandle;

    private ObjectFactory(Class<T> clazz, MethodHandle constructorHandle) {
        this.clazz = clazz;
        this.constructorHandle = constructorHandle;
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
        final MethodHandles.Lookup lookup = MethodHandles.lookup();

        final MethodHandle constructorHandle;
        try {
            constructorHandle = lookup.findConstructor(clazz, MethodType.methodType(void.class));
        } catch (NoSuchMethodException | IllegalAccessException e) {
            throw new ObjectCreationException("failed to find constructor of '" + clazz + "'", e);
        }

        return new ObjectFactory<>(clazz, constructorHandle);
    }

    /**
     * Creates a new instance of the object
     *
     * @return a new instance of the object
     * @throws ObjectCreationException if failed to create the object
     */
    @SuppressWarnings("unchecked")
    @NotNull
    public T newInstance() {
        try {
            return (T) constructorHandle.invoke();
        } catch (ClassCastException ce) {
            throw new AssertionError("Shouldn't happen", ce);
        } catch (Throwable e) {
            throw new ObjectCreationException("failed to create new instance of '" + clazz + "'", e);
        }
    }

}
