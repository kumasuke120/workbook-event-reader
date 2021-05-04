package com.github.kumasuke120.excel;


import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Collections;
import java.util.Objects;
import java.util.function.Function;

import static com.github.kumasuke120.excel.WorkbookDateTimeFormatters.DEFAULT_DATE_TIME_FORMATTERS;
import static com.github.kumasuke120.excel.WorkbookDateTimeFormatters.parseTemporalAccessor;

/**
 * A value class which encapsulates all possible values that the {@link WorkbookEventReader} may return and which
 * provides convenient ways to convert between them
 */
public final class CellValue {

    /**
     * A <code>CellValue</code> singleton whose value is <code>null</code>
     */
    private static final CellValue NULL = new CellValue(null);

    private final Object originalValue;

    private CellValue(@Nullable Object originalValue) {
        assert isTypeAllowed(originalValue);

        this.originalValue = originalValue;
    }

    /**
     * Returns a <code>CellValue</code> based on given value.<br>
     * If the given value is <code>null</code>, it will always return the same instance.
     *
     * @param originalValue the given value
     * @return an instance of <code>CellValue</code>
     */
    @NotNull
    @Contract("!null -> new")
    static CellValue newInstance(@Nullable Object originalValue) {
        if (originalValue == null) {
            return NULL;
        } else {
            return new CellValue(originalValue);
        }
    }

    private boolean isTypeAllowed(@Nullable Object originalValue) {
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
     * Returns the original value of this {@link CellValue}.<br>
     * The original value will be one of the following types:
     * <ul>
     * <li><code>null</code></li>
     * <li>{@link Boolean}</li>
     * <li>{@link Integer}</li>
     * <li>{@link Long}</li>
     * <li>{@link Double}</li>
     * <li>{@link String}</li>
     * <li>{@link LocalTime}</li>
     * <li>{@link LocalDate}</li>
     * <li>{@link LocalDateTime}</li>
     * </ul>
     *
     * @return original value
     */
    @Nullable
    public Object originalValue() {
        return originalValue;
    }

    /**
     * Returns type of the original value. It will return <code>null</code> if the original value is <code>null</code>.
     *
     * @return the type of the original value if possible, otherwise a {@link NullPointerException} will be thrown
     * @throws NullPointerException the original value is null
     */
    @NotNull
    public Class<?> originalType() {
        return Objects.requireNonNull(originalValue).getClass();
    }

    /**
     * Checks if the original value is <code>null</code>.
     *
     * @return <code>true</code> if the original value is <code>null</code>, otherwise <code>false</code>
     */
    public boolean isNull() {
        return originalValue == null;
    }

    /**
     * Converts the original value to its <code>boolean</code> counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type <code>boolean</code>, it will
     * be returned directly.
     *
     * @return <code>boolean</code> version of the original value
     * @throws CellValueCastException cannot convert the original value to <code>boolean</code> type
     */
    public boolean booleanValue() {
        if (originalValue instanceof Boolean) {
            return (Boolean) originalValue;
        } else if (originalValue instanceof Number) {
            return ((Number) originalValue).intValue() != 0;
        } else {
            throw new CellValueCastException();
        }
    }

    /**
     * Converts the original value to its <code>int</code> counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type <code>int</code>, it will
     * be returned directly.
     *
     * @return <code>int</code> version of the original value
     * @throws CellValueCastException cannot convert the original value to <code>int</code> type
     */
    public int intValue() {
        if (originalValue instanceof Integer) {
            return (Integer) originalValue;
        } else if (originalValue instanceof Number) {
            return ((Number) originalValue).intValue();
        } else if (originalValue instanceof String) {
            final String originalString = (String) this.originalValue;

            try {
                return Integer.parseInt(originalString);
            } catch (NumberFormatException e1) {
                try {
                    return (int) Double.parseDouble(originalString);
                } catch (NumberFormatException e2) {
                    // it will be more reasonable to throw an exception which
                    // indicates that the conversion to int has failed
                    e1.addSuppressed(e2);
                    throw new CellValueCastException(e1);
                }
            }
        } else {
            throw new CellValueCastException();
        }
    }

    /**
     * Converts the original value to its <code>long</code> counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type <code>long</code>, it will
     * be returned directly.
     *
     * @return <code>long</code> version of the original value
     * @throws CellValueCastException cannot convert the original value to <code>long</code> type
     */
    public long longValue() {
        if (originalValue instanceof Long) {
            return (Long) originalValue;
        } else if (originalValue instanceof Number) {
            return ((Number) originalValue).longValue();
        } else if (originalValue instanceof String) {
            final String originalString = (String) this.originalValue;

            try {
                return Long.parseLong(originalString);
            } catch (NumberFormatException e1) {
                try {
                    return (long) Double.parseDouble(originalString);
                } catch (NumberFormatException e2) {
                    // it will be more reasonable to throw an exception which
                    // indicates that the conversion to long has failed
                    e1.addSuppressed(e2);
                    throw new CellValueCastException(e1);
                }
            }
        } else {
            throw new CellValueCastException();
        }
    }

    /**
     * Converts the original value to its <code>double</code> counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type <code>double</code>, it will
     * be returned directly.
     *
     * @return <code>double</code> version of the original value
     * @throws CellValueCastException cannot convert the original value to <code>double</code> type
     */
    public double doubleValue() {
        if (originalValue instanceof Double) {
            return (Double) originalValue;
        } else if (originalValue instanceof Number) {
            return ((Number) originalValue).doubleValue();
        } else if (originalValue instanceof String) {
            try {
                return Double.parseDouble((String) originalValue);
            } catch (NumberFormatException e) {
                throw new CellValueCastException(e);
            }
        } else {
            throw new CellValueCastException();
        }
    }

    /**
     * Converts the original value to its {@link String} counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type {@link String}, it will
     * be returned directly. Otherwise, the original value will be converted to {@link String} via
     * {@link String#valueOf(Object)}.
     *
     * @return {@link String} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     */
    @NotNull
    public String stringValue() {
        if (originalValue != null) {
            return String.valueOf(originalValue);
        } else {
            // indicates the originalValue is null
            throw new CellValueCastException(new NullPointerException());
        }
    }

    /**
     * Converts the original value to its {@link LocalTime} counterpart.<br>
     * It will use all preset {@link DateTimeFormatter} static fields that is of type {@link DateTimeFormatter}
     * if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalTime}, it will
     * be returned directly.
     *
     * @return {@link LocalTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     */
    @NotNull
    public LocalTime localTimeValue() {
        return localTimeValue(DEFAULT_DATE_TIME_FORMATTERS);
    }

    /**
     * Converts the original value to its {@link LocalTime} counterpart.<br>
     * The given formatter will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalTime}, it will
     * be returned directly.
     *
     * @param formatter {@link DateTimeFormatter} instance to help converting {@link String} to {@link LocalTime}
     * @return {@link LocalTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatter is <code>null</code>
     */
    @NotNull
    public LocalTime localTimeValue(@NotNull(exception = NullPointerException.class) DateTimeFormatter formatter) {
        Objects.requireNonNull(formatter);
        return localTimeValue(Collections.singleton(formatter));
    }

    /**
     * Converts the original value to its {@link LocalTime} counterpart.<br>
     * The given formatters will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalTime}, it will
     * be returned directly.
     *
     * @param formatters {@link DateTimeFormatter} instances to help converting {@link String} to {@link LocalTime}
     * @return {@link LocalTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatters is <code>null</code>
     */
    @NotNull
    public LocalTime localTimeValue(@NotNull(exception = NullPointerException.class)
                                            Iterable<DateTimeFormatter> formatters) {
        Objects.requireNonNull(formatters);

        if (originalValue instanceof LocalTime) {
            return (LocalTime) originalValue;
        } else if (originalValue instanceof LocalDateTime) {
            return ((LocalDateTime) originalValue).toLocalTime();
        } else if (originalValue instanceof String) {
            final TemporalAccessor temporalAccessor = parseTemporalAccessor((String) originalValue, formatters);

            try {
                return LocalTime.from(temporalAccessor);
            } catch (DateTimeException e1) {
                try {
                    return LocalDateTime.from(temporalAccessor).toLocalTime();
                } catch (DateTimeException e2) {
                    // it will be more reasonable to throw an exception which
                    // indicates that the conversion to LocalTime has failed
                    e1.addSuppressed(e2);
                    throw new CellValueCastException(e1);
                }
            }
        } else {
            throw new CellValueCastException();
        }
    }

    /**
     * Converts the original value to its {@link LocalDate} counterpart.<br>
     * It will use all preset {@link DateTimeFormatter} static fields that is of type {@link DateTimeFormatter}
     * if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDate}, it will
     * be returned directly.
     *
     * @return {@link LocalDate} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     */
    @NotNull
    public LocalDate localDateValue() {
        return localDateValue(DEFAULT_DATE_TIME_FORMATTERS);
    }

    /**
     * Converts the original value to its {@link LocalDate} counterpart.<br>
     * The given formatter will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDate}, it will
     * be returned directly.
     *
     * @param formatter {@link DateTimeFormatter} instance to help converting {@link String} to {@link LocalDate}
     * @return {@link LocalDate} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatter is <code>null</code>
     */
    @NotNull
    public LocalDate localDateValue(@NotNull(exception = NullPointerException.class)
                                            DateTimeFormatter formatter) {
        Objects.requireNonNull(formatter);
        return localDateValue(Collections.singleton(formatter));
    }

    /**
     * Converts the original value to its {@link LocalDate} counterpart.<br>
     * The given formatters will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDate}, it will
     * be returned directly.
     *
     * @param formatters {@link DateTimeFormatter} instances to help converting {@link String} to {@link LocalDate}
     * @return {@link LocalDate} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatters is <code>null</code>
     */
    @NotNull
    public LocalDate localDateValue(@NotNull(exception = NullPointerException.class)
                                            Iterable<DateTimeFormatter> formatters) {
        Objects.requireNonNull(formatters);

        if (originalValue instanceof LocalDate) {
            return (LocalDate) originalValue;
        } else if (originalValue instanceof LocalDateTime) {
            return ((LocalDateTime) originalValue).toLocalDate();
        } else if (originalValue instanceof String) {
            try {
                return LocalDate.from(parseTemporalAccessor((String) originalValue, formatters));
            } catch (DateTimeException e) {
                throw new CellValueCastException(e);
            }
        } else {
            throw new CellValueCastException();
        }
    }

    /**
     * Converts the original value to its {@link LocalDateTime} counterpart.<br>
     * It will use all preset {@link DateTimeFormatter} static fields that is of type {@link DateTimeFormatter}
     * if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDateTime}, it will
     * be returned directly.
     *
     * @return {@link LocalDateTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     */
    @NotNull
    public LocalDateTime localDateTimeValue() {
        return localDateTimeValue(DEFAULT_DATE_TIME_FORMATTERS);
    }

    /**
     * Converts the original value to its {@link LocalDateTime} counterpart.<br>
     * The given formatter will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDateTime}, it will
     * be returned directly.
     *
     * @param formatter {@link DateTimeFormatter} instance to help converting {@link String} to {@link LocalDateTime}
     * @return {@link LocalDateTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatter is <code>null</code>
     */
    @NotNull
    public LocalDateTime localDateTimeValue(@NotNull(exception = NullPointerException.class)
                                                    DateTimeFormatter formatter) {
        Objects.requireNonNull(formatter);
        return localDateTimeValue(Collections.singleton(formatter));
    }

    /**
     * Converts the original value to its {@link LocalDateTime} counterpart.<br>
     * The given formatters will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDateTime}, it will
     * be returned directly.
     *
     * @param formatters {@link DateTimeFormatter} instances to help converting {@link String} to {@link LocalDateTime}
     * @return {@link LocalDateTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatters is <code>null</code>
     */
    @NotNull
    public LocalDateTime localDateTimeValue(@NotNull(exception = NullPointerException.class)
                                                    Iterable<DateTimeFormatter> formatters) {
        Objects.requireNonNull(formatters);

        if (originalValue instanceof LocalDateTime) {
            return (LocalDateTime) originalValue;
        } else if (originalValue instanceof LocalDate) {
            return ((LocalDate) originalValue).atStartOfDay();
        } else if (originalValue instanceof String) {
            final TemporalAccessor temporalAccessor = parseTemporalAccessor((String) originalValue, formatters);

            try {
                return LocalDateTime.from(temporalAccessor);
            } catch (DateTimeException e1) {
                try {
                    return LocalDate.from(temporalAccessor).atStartOfDay();
                } catch (DateTimeException e2) {
                    // it will be more reasonable to throw an exception which
                    // indicates that the conversion to LocalDateTime has failed
                    e1.addSuppressed(e2);
                    throw new CellValueCastException(e1);
                }
            }
        } else {
            throw new CellValueCastException();
        }
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
    @NotNull
    @Contract(pure = true)
    public CellValue mapOriginalValue(@NotNull(exception = NullPointerException.class)
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

            return CellValue.newInstance(newOriginalValue);
        }
    }

    /**
     * Trims the original value if its type is of {@link String}.<br>
     * A new instance of {@link CellValue} will be created if there is any change to this
     * {@link CellValue} instance.
     *
     * @return {@link CellValue} based on the trimmed value
     */
    @NotNull
    @Contract(pure = true)
    public CellValue trim() {
        return mapOriginalValue(v -> v instanceof String ? ((String) v).trim() : v);
    }

    @Override
    public boolean equals(@Nullable Object o) {
        if (this == o) return true;
        if (!(o instanceof CellValue)) return false;
        CellValue cellValue = (CellValue) o;
        return Objects.equals(originalValue, cellValue.originalValue);
    }

    @Override
    public int hashCode() {
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
     * The form aforementioned may not be holden in future. One should not
     * use its value to extract information.
     *
     * @return a string representation of the <code>CellValue</code>
     */
    @Override
    @NotNull
    public String toString() {
        return getClass().getName() +
                "{" +
                (isNull() ? "" : "type = " + originalType().getCanonicalName() + ", ") +
                "value = " + originalValue() +
                "}" +
                "@" + Integer.toHexString(System.identityHashCode(this));
    }


}
