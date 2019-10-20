package app.kumasuke.excel;


import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.TemporalAccessor;
import java.util.Collections;
import java.util.HashSet;
import java.util.Objects;
import java.util.Set;

/**
 * A value class which encapsulates all possible values that the {@link WorkbookEventReader} may return and which
 * provides convenient ways to convert between them
 */
@SuppressWarnings("WeakerAccess")
public final class CellValue {
    /**
     * A <code>CellValue</code> singleton whose value is <code>null</code>
     */
    private static final CellValue NULL = new CellValue(null);

    /**
     * The default <code>DateTimeFormatter</code>s that will be used when converting <code>String</code>
     * to <code>LocalTime</code>, <code>LocalDate</code>, <code>LocalDateTime</code>
     */
    private static final Set<DateTimeFormatter> DEFAULT_DATE_TIME_FORMATTERS;

    static {
        final Set<DateTimeFormatter> formatters = new HashSet<>();
        final Field[] fields = DateTimeFormatter.class.getFields();
        for (Field field : fields) {
            final int modifiers = field.getModifiers();

            // public static final DateTimeFormatter ...
            if (Modifier.isStatic(modifiers) && Modifier.isFinal(modifiers) &&
                    DateTimeFormatter.class.equals(field.getType())) {
                try {
                    final DateTimeFormatter f = (DateTimeFormatter) field.get(null);
                    formatters.add(f);
                } catch (IllegalAccessException e) {
                    throw new AssertionError(e);
                }
            }
        }

        DEFAULT_DATE_TIME_FORMATTERS = Collections.unmodifiableSet(formatters);
    }

    private final Object originalValue;

    private CellValue(Object originalValue) {
        assert originalValue == null ||
                originalValue instanceof Boolean ||
                originalValue instanceof Integer ||
                originalValue instanceof Long ||
                originalValue instanceof Double ||
                originalValue instanceof String ||
                originalValue instanceof LocalTime ||
                originalValue instanceof LocalDate ||
                originalValue instanceof LocalDateTime;

        this.originalValue = originalValue;
    }

    /**
     * Returns a <code>CellValue</code> based on given value.<br>
     * If the given value is <code>null</code>, it will always return the same instance.
     *
     * @param originalValue the given value
     * @return an instance of <code>CellValue</code>
     */
    static CellValue newInstance(Object originalValue) {
        if (originalValue == null) {
            return NULL;
        } else {
            return new CellValue(originalValue);
        }
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
    public Object originalValue() {
        return originalValue;
    }

    /**
     * Returns type of the original value. It will return <code>null</code> if the original value is <code>null</code>.
     *
     * @return the type of the original value if possible, otherwise a {@link NullPointerException} will be thrown
     * @throws NullPointerException the original value is null
     */
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
    public LocalTime localTimeValue(DateTimeFormatter formatter) {
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
    public LocalTime localTimeValue(Iterable<DateTimeFormatter> formatters) {
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
    public LocalDate localDateValue(DateTimeFormatter formatter) {
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
    public LocalDate localDateValue(Iterable<DateTimeFormatter> formatters) {
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
    public LocalDateTime localDateTimeValue(DateTimeFormatter formatter) {
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
    public LocalDateTime localDateTimeValue(Iterable<DateTimeFormatter> formatters) {
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

    private TemporalAccessor parseTemporalAccessor(String value, Iterable<DateTimeFormatter> formatters) {
        RuntimeException previousEx = null;

        for (DateTimeFormatter formatter : formatters) {
            if (formatter == null) {
                if (previousEx == null) {
                    previousEx = new NullPointerException();
                } else {
                    previousEx.addSuppressed(new NullPointerException());
                }
                continue;
            }

            try {
                return formatter.parse(value);
            } catch (DateTimeParseException e) {
                if (previousEx != null) {
                    previousEx.addSuppressed(e);
                }
                previousEx = e;
            }
        }

        assert previousEx != null;
        throw new CellValueCastException(previousEx);
    }

    @Override
    public boolean equals(Object o) {
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
    public String toString() {
        return getClass().getName() +
                "{" +
                (isNull() ? "" : "type = " + originalType().getCanonicalName() + ", ") +
                "value = " + originalValue() +
                "}" +
                "@" + Integer.toHexString(System.identityHashCode(this));
    }
}
