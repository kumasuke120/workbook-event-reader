package com.github.kumasuke120.excel.util;

import org.jetbrains.annotations.ApiStatus;

import java.util.Collection;

/**
 * A utility class that provides some common operations on {@link Collection}
 */
@ApiStatus.Internal
public class CollectionUtils {

    private CollectionUtils() {
        throw new UnsupportedOperationException();
    }

    /**
     * Checks if the given {@link Collection} is {@code null} or empty
     *
     * @param coll the {@link Collection} to be checked
     * @return {@code true} if the given {@link Collection} is {@code null} or empty, otherwise {@code false}
     */
    public static boolean isEmpty(Collection<?> coll) {
        return coll == null || coll.isEmpty();
    }

    /**
     * Checks if the given {@link Collection} is not {@code null} and not empty
     *
     * @param coll the {@link Collection} to be checked
     * @return {@code true} if the given {@link Collection} is not {@code null} and not empty, otherwise {@code false}
     */
    public static boolean isNotEmpty(Collection<?> coll) {
        return !isEmpty(coll);
    }

}
