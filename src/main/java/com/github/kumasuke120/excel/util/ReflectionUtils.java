package com.github.kumasuke120.excel.util;

import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.Nullable;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;

/**
 * A utility class for reflection-related operations.
 * This class provides methods to access classes and methods using reflection.
 */
@ApiStatus.Internal
public class ReflectionUtils {

    private static final MethodHandles.Lookup handleLookup = MethodHandles.publicLookup();

    private ReflectionUtils() {
        throw new UnsupportedOperationException();
    }

    /**
     * Gets the class object for the specified class name.
     *
     * @param className the name of the class
     * @return the class object, or null if the class cannot be found
     */
    @Nullable
    public static Class<?> getClass(String className) {
        try {
            return Class.forName(className);
        } catch (ClassNotFoundException e) {
            return null;
        }
    }

    /**
     * Creates a new instance of the specified class using the default constructor.
     *
     * @param clazz the class
     * @param name  the name of the static field
     * @param type  the type of the static field
     * @return the static field <code>MethodHandle</code>, or null if not found
     */
    @Nullable
    public static MethodHandle findStaticGetterHandle(Class<?> clazz, String name, Class<?> type) {
        try {
            return handleLookup.findStaticGetter(clazz, name, type);
        } catch (IllegalAccessException | NoSuchFieldException e) {
            return null;
        }
    }

    /**
     * Finds a static method handle for the specified class, method name, and method type.
     *
     * @param clazz the class
     * @param name  the method name
     * @param type  the method type
     * @return the static <code>MethodHandle</code>, or null if not found
     */
    @Nullable
    public static MethodHandle findStaticHandle(Class<?> clazz, String name, MethodType type) {
        try {
            return handleLookup.findStatic(clazz, name, type);
        } catch (NoSuchMethodException | IllegalAccessException e) {
            return null;
        }
    }

    /**
     * Finds a virtual method handle for the specified class, method name, and method type.
     *
     * @param clazz the class
     * @param name  the method name
     * @param type  the method type
     * @return the virtual <code>MethodHandle</code>, or null if not found
     */
    @Nullable
    public static MethodHandle findVirtualHandle(Class<?> clazz, String name, MethodType type) {
        try {
            return handleLookup.findVirtual(clazz, name, type);
        } catch (NoSuchMethodException | IllegalAccessException e) {
            return null;
        }
    }

}
