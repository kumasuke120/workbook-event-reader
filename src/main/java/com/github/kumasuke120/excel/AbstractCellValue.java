package com.github.kumasuke120.excel;

import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Objects;
import java.util.function.Function;

/**
 * The base class for {@link CellValue}, containing common methods and utilities
 */
public abstract class AbstractCellValue implements CellValue {

    final Object originalValue;

    /**
     * Creates a new instance of {@link CellValue} based on the given value
     *
     * @param originalValue the given value
     */
    AbstractCellValue(Object originalValue) {
        assert isTypeAllowed(originalValue);

        this.originalValue = originalValue;
    }

    /**
     * Returns a <code>CellValue</code> based on the given value.<br>
     * If the given value is <code>null</code>, it will always return the same instance.
     *
     * @param originalValue the given value
     * @return an instance of <code>CellValue</code>
     */
    @NotNull
    abstract CellValue valueOf(@Nullable Object originalValue);

    /**
     * Determines whether the given value could be constructed as a {@link CellValue}.
     *
     * @param originalValue the given original value
     * @return is the original value allowed
     */
    boolean isTypeAllowed(@Nullable Object originalValue) {
        return originalValue == null ||
                originalValue instanceof Boolean ||
                originalValue instanceof Integer ||
                originalValue instanceof Long ||
                originalValue instanceof Double ||
                originalValue instanceof String ||
                originalValue instanceof LocalTime ||
                originalValue instanceof LocalDate ||
                originalValue instanceof LocalDateTime;
    }

    /**
     * Returns type of the original value. It will return <code>null</code> if the original value is <code>null</code>.
     *
     * @return the type of the original value if possible, otherwise a {@link NullPointerException} will be thrown
     * @throws NullPointerException the original value is null
     */
    @Override
    @NotNull
    public final Class<?> originalType() {
        return Objects.requireNonNull(originalValue).getClass();
    }


    /**
     * Returns the original value of this {@link CellValue}.
     *
     * @return original value
     */
    @Override
    @Nullable
    public final Object originalValue() {
        return originalValue;
    }


    /**
     * Checks if the original value is <code>null</code>.
     *
     * @return <code>true</code> if the original value is <code>null</code>, otherwise <code>false</code>
     */
    @Override
    public final boolean isNull() {
        return originalValue == null;
    }

    /**
     * Uses the given function to map the {@link #originalValue()} to a new value, and returns a new
     * {@link CellValue} based on the new value.<br>
     * The return value of <code>mappingFunction</code> must be one of the types listed on the javadoc of
     * {@link #originalValue()}.
     *
     * @param mappingFunction value mapping function
     * @return {@link CellValue} based on the mapped value
     * @throws CellValueCastException cannot perform mapping or the mapped value is not in the allow type
     */
    @Override
    @NotNull
    @Contract(pure = true)
    public final CellValue mapOriginalValue(@NotNull(exception = NullPointerException.class)
                                      Function<Object, Object> mappingFunction) {
        Objects.requireNonNull(mappingFunction);

        final Object newOriginalValue;
        try {
            newOriginalValue = mappingFunction.apply(originalValue);
        } catch (RuntimeException e) {
            throw new CellValueCastException(e);
        }

        if (newOriginalValue == originalValue) { // prevents creating a new instance if two values are identical
            return this;
        } else {
            if (!isTypeAllowed(newOriginalValue)) {
                throw new CellValueCastException("mapped value should be in one of the allowed types");
            }

            return valueOf(newOriginalValue);
        }
    }

    /**
     * Trims the original value if its type is of {@link String}.<br>
     * A new instance of {@link CellValue} will be created if there is any change to this
     * {@link CellValue} instance.
     *
     * @return {@link CellValue} based on the trimmed value
     */
    @Override
    @NotNull
    @Contract(pure = true)
    public final CellValue trim() {
        return mapOriginalValue(v -> v instanceof String ? ((String) v).trim() : v);
    }

    @Override
    public final boolean equals(@Nullable Object o) {
        if (this == o) return true;
        if (!(o instanceof AbstractCellValue)) return false;
        AbstractCellValue cellValue = (AbstractCellValue) o;
        return Objects.equals(originalValue, cellValue.originalValue);
    }

    @Override
    public final int hashCode() {
        return Objects.hashCode(originalValue);
    }

    /**
     * Returns the string representation of the <code>CellValue</code>.<br>
     * This result string will follow this form:
     * <pre><code>
     * getClass().getName() +
     * "{" +
     * (isNull() ? "" : "type = <i>value type</i>, ") +
     * "value = <i>value</i>" +
     * "}" +
     * "@" + Integer.toHexString(System.identityHashCode(this))
     * </code></pre>
     * The form aforementioned may not be holden in the future. One should not
     * use its value to extract information.
     *
     * @return a string representation of the <code>CellValue</code>
     */
    @Override
    @NotNull
    public final String toString() {
        return getClass().getName() +
                "{" +
                (isNull() ? "" : "type = " + originalType().getCanonicalName() + ", ") +
                "value = " + originalValue() +
                "}" +
                "@" + Integer.toHexString(System.identityHashCode(this));
    }

}
