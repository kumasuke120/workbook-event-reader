package com.github.kumasuke120.excel;

import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.Objects;
import java.util.function.Function;

import static com.github.kumasuke120.excel.WorkbookDateTimeFormatters.DEFAULT_DATE_TIME_FORMATTERS;

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
    final boolean isTypeAllowed(@Nullable Object originalValue) {
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

    @Override
    @NotNull
    public final Class<?> originalType() {
        return Objects.requireNonNull(originalValue).getClass();
    }

    @Override
    @Nullable
    public final Object originalValue() {
        return originalValue;
    }

    @Override
    public final boolean isNull() {
        return originalValue == null;
    }

    @Override
    @NotNull
    public final LocalTime localTimeValue() {
        return localTimeValue(DEFAULT_DATE_TIME_FORMATTERS);
    }

    @Override
    @NotNull
    public final LocalTime localTimeValue(@NotNull(exception = NullPointerException.class) DateTimeFormatter formatter) {
        Objects.requireNonNull(formatter);
        return localTimeValue(Collections.singleton(formatter));
    }

    @Override
    @NotNull
    public final LocalDate localDateValue() {
        return localDateValue(DEFAULT_DATE_TIME_FORMATTERS);
    }

    @Override
    @NotNull
    public final LocalDate localDateValue(@NotNull(exception = NullPointerException.class)
                                    DateTimeFormatter formatter) {
        Objects.requireNonNull(formatter);
        return localDateValue(Collections.singleton(formatter));
    }

    @Override
    @NotNull
    public final LocalDateTime localDateTimeValue() {
        return localDateTimeValue(DEFAULT_DATE_TIME_FORMATTERS);
    }

    @Override
    @NotNull
    public final LocalDateTime localDateTimeValue(@NotNull(exception = NullPointerException.class)
                                            DateTimeFormatter formatter) {
        Objects.requireNonNull(formatter);
        return localDateTimeValue(Collections.singleton(formatter));
    }

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
