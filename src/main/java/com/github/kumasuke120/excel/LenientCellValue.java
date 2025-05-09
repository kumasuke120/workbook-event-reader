package com.github.kumasuke120.excel;


import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.math.BigDecimal;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Objects;

import static com.github.kumasuke120.excel.WorkbookDateTimeFormatters.parseTemporalAccessor;

/**
 * A value class which encapsulates all possible values that the {@link WorkbookEventReader} may return and which
 * provides convenient and lenient ways to convert between them
 */
@ApiStatus.Internal
final class LenientCellValue extends AbstractCellValue {

    /**
     * A <code>CellValue</code> singleton whose value is <code>null</code>
     */
    private static final LenientCellValue NULL = new LenientCellValue(null);

    private LenientCellValue(@Nullable Object originalValue) {
        super(originalValue);
    }

    /**
     * Returns a <code>LenientCellValue</code> based on the given value.<br>
     * If the given value is <code>null</code>, it will always return the same instance.
     *
     * @param originalValue the given value
     * @return an instance of <code>CellValue</code>
     */
    @NotNull
    @Contract("!null -> new")
    static LenientCellValue newInstance(@Nullable Object originalValue) {
        if (originalValue == null) {
            return NULL;
        } else {
            return new LenientCellValue(originalValue);
        }
    }

    @Override
    @NotNull
    CellValue valueOf(@Nullable Object originalValue) {
        return newInstance(originalValue);
    }

    @Override
    public boolean booleanValue() {
        if (originalValue instanceof Boolean) {
            return (Boolean) originalValue;
        } else if (originalValue instanceof Number) {
            return ((Number) originalValue).intValue() != 0;
        } else if (originalValue instanceof String) {
            final String originalString = (String) this.originalValue;

            if (originalString.equalsIgnoreCase("true")) {
                return true;
            } else if (originalString.equalsIgnoreCase("false")) {
                return false;
            } else {
                try {
                    return Integer.parseInt(originalString) != 0;
                } catch (NumberFormatException e1) {
                    try {
                        return ((int) Double.parseDouble(originalString)) != 0;
                    } catch (NumberFormatException e2) {
                        // it will be more reasonable to throw an exception which
                        // indicates that the conversion to boolean has failed
                        e1.addSuppressed(e2);
                        throw new CellValueCastException(e1);
                    }
                }
            }
        } else {
            throw new CellValueCastException();
        }
    }

    @Override
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

    @Override
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

    @Override
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

    @Override
    public BigDecimal bigDecimalValue() {
        if (originalValue instanceof BigDecimal) {
            return (BigDecimal) originalValue;
        } else if (originalValue instanceof Number) {
            return new BigDecimal(originalValue.toString());
        } else if (originalValue instanceof String) {
            try {
                return new BigDecimal((String) originalValue);
            } catch (NumberFormatException e) {
                throw new CellValueCastException(e);
            }
        } else {
            throw new CellValueCastException();
        }
    }

    @Override
    @NotNull
    public String stringValue() {
        if (originalValue != null) {
            return String.valueOf(originalValue);
        } else {
            // indicates the originalValue is null
            throw new CellValueCastException(new NullPointerException());
        }
    }

    @Override
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
     * The given formatters will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDate}, it will
     * be returned directly.
     *
     * @param formatters {@link DateTimeFormatter} instances to help converting {@link String} to {@link LocalDate}
     * @return {@link LocalDate} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatters is <code>null</code>
     */
    @Override
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

    @Override
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

}
