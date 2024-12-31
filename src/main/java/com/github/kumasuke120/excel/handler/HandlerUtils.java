package com.github.kumasuke120.excel.handler;

import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.*;
import java.util.Date;

/**
 * A utility class for handling various operations related to workbook records.
 * This class contains methods for ensuring a class is annotated with {@code @WorkbookRecord},
 * retrieving default values for primitive types, and converting values to different types.
 * <p>
 * This class is intended for internal use only.
 * </p>
 */
@ApiStatus.Internal
class HandlerUtils {

    private HandlerUtils() {
        throw new UnsupportedOperationException();
    }

    /**
     * Ensures that the specified class is annotated with {@code @WorkbookRecord}.
     *
     * @param clazz the class to check
     * @param <T>   the type of the given class
     * @return the specified class
     * @throws WorkbookRecordException if the specified class is not annotated with {@code @WorkbookRecord}
     */
    static <T> Class<T> ensureWorkbookRecordClass(@Nullable Class<T> clazz) {
        if (clazz == null) {
            throw new NullPointerException();
        }

        WorkbookRecord a = clazz.getAnnotation(WorkbookRecord.class);
        if (a == null) {
            throw new WorkbookRecordException("cannot handle the class without being annotated by @WorkbookRecord");
        }

        return clazz;
    }

    /**
     * Retrieves the default value for the specified class.
     *
     * @param clazz the given class
     * @return the default value for the specified class
     */
    static Object getDefaultValue(@NotNull Class<?> clazz) {
        if (!clazz.isPrimitive()) {
            return null;
        }

        if (int.class == clazz) {
            return 0;
        } else if (byte.class == clazz) {
            return (byte) 0;
        } else if (short.class == clazz) {
            return (short) 0;
        } else if (long.class == clazz) {
            return 0L;
        } else if (float.class == clazz) {
            return 0.0f;
        } else if (double.class == clazz) {
            return 0.0d;
        } else if (boolean.class == clazz) {
            return false;
        } else if (char.class == clazz) {
            return '\0';
        } else {
            throw new AssertionError("Shouldn't happen");
        }
    }

    /**
     * Converts the specified value to a boolean as much as possible.
     *
     * @param value the value to convert
     * @return the converted boolean value
     * @throws IllegalArgumentException if the value cannot be converted to a boolean
     */
    static boolean asBoolean(@NotNull Object value) {
        if (value instanceof Boolean) {
            return (boolean) value;
        } else if (value instanceof Number) {
            return ((Number) value).intValue() != 0;
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to boolean");
        }
    }

    /**
     * Converts the specified value to an integer as much as possible.
     *
     * @param value the value to convert
     * @return the converted integer value
     * @throws IllegalArgumentException if the value cannot be converted to an integer
     */
    static byte asByte(@NotNull Object value) {
        if (value instanceof Byte) {
            return (byte) value;
        } else if (value instanceof Number) {
            return ((Number) value).byteValue();
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to byte");
        }
    }

    /**
     * Converts the specified value to a short as much as possible.
     *
     * @param value the value to convert
     * @return the converted short value
     * @throws IllegalArgumentException if the value cannot be converted to a short
     */
    static short asShort(@NotNull Object value) {
        if (value instanceof Short) {
            return (short) value;
        } else if (value instanceof Number) {
            return ((Number) value).shortValue();
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to short");
        }
    }

    /**
     * Converts the specified value to a float as much as possible.
     *
     * @param value the value to convert
     * @return the converted float value
     * @throws IllegalArgumentException if the value cannot be converted to a float
     */
    static float asFloat(@NotNull Object value) {
        if (value instanceof Float) {
            return (float) value;
        } else if (value instanceof Number) {
            return ((Number) value).floatValue();
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to float");
        }
    }

    /**
     * Converts the specified value to a {@link BigInteger} as much as possible.
     *
     * @param value the value to convert
     * @return the converted {@link BigInteger} value
     * @throws IllegalArgumentException if the value cannot be converted to a {@link BigInteger}
     */
    static BigInteger asBigInteger(@NotNull Object value) {
        if (value instanceof BigInteger) {
            return (BigInteger) value;
        } else if (value instanceof BigDecimal) {
            return ((BigDecimal) value).toBigInteger();
        } else if (value instanceof Number) {
            return BigInteger.valueOf(((Number) value).longValue());
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to Date");
        }
    }

    /**
     * Converts the specified value to a {@link Date} as much as possible.
     *
     * @param value the value to convert
     * @return the converted {@link Date} value
     * @throws IllegalArgumentException if the value cannot be converted to a {@link Date}
     */
    @NotNull
    static Date asDate(@NotNull Object value) {
        if (value instanceof Date) {
            return (Date) value;
        } else if (value instanceof LocalDate) {
            final LocalDate localDateValue = (LocalDate) value;
            final Instant instant = localDateValue.atStartOfDay(ZoneId.systemDefault()).toInstant();
            return Date.from(instant);
        } else if (value instanceof LocalDateTime) {
            final LocalDateTime localDateTimeValue = (LocalDateTime) value;
            final Instant instant = localDateTimeValue.atZone(ZoneId.systemDefault()).toInstant();
            return Date.from(instant);
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to Date");
        }
    }

    /**
     * Converts the specified value to a {@link java.sql.Time} as much as possible.
     *
     * @param value the value to convert
     * @return the converted {@link java.sql.Time} value
     * @throws IllegalArgumentException if the value cannot be converted to a {@link java.sql.Time}
     */
    @NotNull
    static java.sql.Time asSqlTime(@NotNull Object value) {
        if (value instanceof java.sql.Time) {
            return (java.sql.Time) value;
        } else if (value instanceof LocalTime) {
            return java.sql.Time.valueOf((LocalTime) value);
        } else {
            try {
                Date date = asDate(value);
                return new java.sql.Time(date.getTime());
            } catch (IllegalArgumentException e) {
                throw new IllegalArgumentException("cannot convert '" + value + "' to Time", e);
            }
        }
    }

    /**
     * Converts the specified value to a {@link java.sql.Date} as much as possible.
     *
     * @param value the value to convert
     * @return the converted {@link java.sql.Date} value
     * @throws IllegalArgumentException if the value cannot be converted to a {@link java.sql.Date}
     */
    @NotNull
    static java.sql.Date asSqlDate(@NotNull Object value) {
        if (value instanceof java.sql.Date) {
            return (java.sql.Date) value;
        } else if (value instanceof LocalDate) {
            return java.sql.Date.valueOf((LocalDate) value);
        } else if (value instanceof LocalDateTime) {
            final LocalDateTime localDateTimeValue = (LocalDateTime) value;
            return java.sql.Date.valueOf(localDateTimeValue.toLocalDate());
        } else {
            try {
                Date date = asDate(value);
                return new java.sql.Date(date.getTime());
            } catch (IllegalArgumentException e) {
                throw new IllegalArgumentException("cannot convert '" + value + "' to Date", e);
            }
        }
    }

    /**
     * Converts the specified value to a {@link java.sql.Timestamp} as much as possible.
     *
     * @param value the value to convert
     * @return the converted {@link java.sql.Timestamp} value
     * @throws IllegalArgumentException if the value cannot be converted to a {@link java.sql.Timestamp}
     */
    @NotNull
    static java.sql.Timestamp asSqlTimestamp(@NotNull Object value) {
        if (value instanceof java.sql.Timestamp) {
            return (java.sql.Timestamp) value;
        } else if (value instanceof LocalDateTime) {
            return java.sql.Timestamp.valueOf((LocalDateTime) value);
        } else if (value instanceof LocalDate) {
            return java.sql.Timestamp.valueOf(((LocalDate) value).atStartOfDay());
        } else {
            try {
                Date date = asDate(value);
                return new java.sql.Timestamp(date.getTime());
            } catch (IllegalArgumentException e) {
                throw new IllegalArgumentException("cannot convert '" + value + "' to Timestamp", e);
            }
        }
    }

}
