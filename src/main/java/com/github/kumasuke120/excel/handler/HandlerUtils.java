package com.github.kumasuke120.excel.handler;

import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.time.*;
import java.util.Date;

@ApiStatus.Internal
class HandlerUtils {

    private HandlerUtils() {
        throw new UnsupportedOperationException();
    }

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

    static boolean asBoolean(@NotNull Object value) {
        if (value instanceof Boolean) {
            return (boolean) value;
        } else if (value instanceof Number) {
            return ((Number) value).intValue() != 0;
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to boolean");
        }
    }

    static byte asByte(@NotNull Object value) {
        if (value instanceof Byte) {
            return (byte) value;
        } else if (value instanceof Number) {
            return ((Number) value).byteValue();
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to byte");
        }
    }

    static short asShort(@NotNull Object value) {
        if (value instanceof Short) {
            return (short) value;
        } else if (value instanceof Number) {
            return ((Number) value).shortValue();
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to short");
        }
    }

    static float asFloat(@NotNull Object value) {
        if (value instanceof Float) {
            return (short) value;
        } else if (value instanceof Number) {
            return ((Number) value).floatValue();
        } else {
            throw new IllegalArgumentException("cannot convert '" + value + "' to float");
        }
    }

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
