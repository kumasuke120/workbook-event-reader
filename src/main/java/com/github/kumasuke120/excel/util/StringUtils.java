package com.github.kumasuke120.excel.util;

import org.jetbrains.annotations.ApiStatus;

/**
 * A utility class that provides some common operations on {@link String}
 */
@ApiStatus.Internal
public class StringUtils {

    private StringUtils() {
        throw new UnsupportedOperationException();
    }

    /**
     * Checks if the given {@link CharSequence} is {@code null} or empty
     *
     * @param charSequence the {@link CharSequence} to be checked
     * @return {@code true} if the given {@link CharSequence} is {@code null} or empty, otherwise {@code false}
     */
    public static boolean isEmpty(CharSequence charSequence) {
        return charSequence == null || charSequence.length() == 0;
    }

    /**
     * Checks if the given {@link CharSequence} is not {@code null} and not empty
     *
     * @param charSequence the {@link CharSequence} to be checked
     * @return {@code true} if the given {@link CharSequence} is not {@code null} and not empty, otherwise {@code false}
     */
    public static boolean isBlank(CharSequence charSequence) {
        if (isEmpty(charSequence)) {
            return true;
        }
        for (int i = 0; i < charSequence.length(); i++) {
            if (!Character.isWhitespace(charSequence.charAt(i))) {
                return false;
            }
        }
        return true;
    }

}
