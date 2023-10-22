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
import java.util.Objects;

import static com.github.kumasuke120.excel.WorkbookDateTimeFormatters.parseTemporalAccessor;

/**
 * A value class which encapsulates all possible values that the {@link WorkbookEventReader} may return and which
 * provides convenient and strict ways to convert between them
 */
public final class StrictCellValue extends AbstractCellValue {

    /**
     * A <code>CellValue</code> singleton whose value is <code>null</code>
     */
    private static final StrictCellValue NULL = new StrictCellValue(null);

    StrictCellValue(Object originalValue) {
        super(originalValue);
    }

    /**
     * Returns a <code>StrictCellValue</code> based on the given value.<br>
     * If the given value is <code>null</code>, it will always return the same instance.
     *
     * @param originalValue the given value
     * @return an instance of <code>CellValue</code>
     */
    @NotNull
    @Contract("!null -> new")
    static StrictCellValue newInstance(@Nullable Object originalValue) {
        if (originalValue == null) {
            return NULL;
        } else {
            return new StrictCellValue(originalValue);
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
        } else {
            throw new CellValueCastException();
        }
    }

    @Override
    public int intValue() {
        if (originalValue instanceof Integer) {
            return (Integer) originalValue;
        } else if (originalValue instanceof String) {
            final String originalString = (String) this.originalValue;

            try {
                return Integer.parseInt(originalString);
            } catch (NumberFormatException e) {
                throw new CellValueCastException(e);
            }
        } else {
            throw new CellValueCastException();
        }
    }

    @Override
    public long longValue() {
        if (originalValue instanceof Long) {
            return (Long) originalValue;
        } else if (originalValue instanceof String) {
            final String originalString = (String) this.originalValue;

            try {
                return Long.parseLong(originalString);
            } catch (NumberFormatException e) {
                throw new CellValueCastException(e);
            }
        } else {
            throw new CellValueCastException();
        }
    }

    @Override
    public double doubleValue() {
        if (originalValue instanceof Double) {
            return (Double) originalValue;
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
    @NotNull
    public String stringValue() {
        if (originalValue instanceof String) {
            return originalValue.toString();
        } else if (originalValue != null) {
            throw new CellValueCastException();
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
        } else if (originalValue instanceof String) {
            final TemporalAccessor temporalAccessor = parseTemporalAccessor((String) originalValue, formatters);

            try {
                return LocalTime.from(temporalAccessor);
            } catch (DateTimeException e) {
                throw new CellValueCastException(e);
            }
        } else {
            throw new CellValueCastException();
        }
    }

    @Override
    public @NotNull LocalDate localDateValue(@NotNull(exception = NullPointerException.class)
                                             Iterable<DateTimeFormatter> formatters) {
        Objects.requireNonNull(formatters);

        if (originalValue instanceof LocalDate) {
            return (LocalDate) originalValue;
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
    public @NotNull LocalDateTime localDateTimeValue(@NotNull(exception = NullPointerException.class)
                                                     Iterable<DateTimeFormatter> formatters) {
        Objects.requireNonNull(formatters);

        if (originalValue instanceof LocalDateTime) {
            return (LocalDateTime) originalValue;
        } else if (originalValue instanceof String) {
            final TemporalAccessor temporalAccessor = parseTemporalAccessor((String) originalValue, formatters);

            try {
                return LocalDateTime.from(temporalAccessor);
            } catch (DateTimeException e) {
                throw new CellValueCastException(e);
            }
        } else {
            throw new CellValueCastException();
        }
    }

}
