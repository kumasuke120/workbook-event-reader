package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.handler.WorkbookRecord.CellValueType;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.*;

@ApiStatus.Internal
class WorkbookRecordProperty<E> {

    static final int COLUMN_NUM_SHEET_INDEX = -1;
    static final int COLUMN_NUM_SHEET_NAME = -2;
    static final int COLUMN_NUM_ROW_NUMBER = -3;

    private final Class<E> recordClass;

    private final int column;

    private final boolean strict;

    private final ValueMethod valueMethod;

    private final Field field;

    WorkbookRecordProperty(@NotNull Class<E> recordClass,
                           int column,
                           boolean strict,
                           @NotNull CellValueType valueType,
                           @NotNull String valueMethodName,
                           @NotNull Field field) {

        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.column = column;
        this.strict = strict;
        this.valueMethod = new ValueMethod(valueType, valueMethodName);
        this.field = field;
    }

    int getColumn() {
        return column;
    }

    void set(@NotNull E record, @NotNull CellValue cellValue) {
        final Object value = valueMethod.getValue(record, cellValue);
        if (!field.isAccessible()) {
            field.setAccessible(true);
        }

        try {
            field.set(record, value);
        } catch (IllegalAccessException e) {
            throw new AssertionError("Shouldn't happen", e);
        } catch (IllegalArgumentException e) {
            throw new WorkbookRecordException("cannot set value on WorkbookRecord", e);
        }
    }

    void set(@NotNull E record, int value) {
        if (!field.isAccessible()) {
            field.setAccessible(true);
        }

        final boolean isPrimitive = field.getType().isPrimitive();

        try {
            if (isPrimitive) {
                field.setInt(record, value);
            } else {
                field.set(record, value);
            }
        } catch (IllegalAccessException e) {
            throw new AssertionError("Shouldn't happen", e);
        } catch (IllegalArgumentException e) {
            throw new WorkbookRecordException("cannot set int value on @WorkbookRecord record class", e);
        }
    }

    void set(@NotNull E record, String value) {
        if (!field.isAccessible()) {
            field.setAccessible(true);
        }

        try {
            field.set(record, value);
        } catch (IllegalAccessException e) {
            throw new AssertionError("Shouldn't happen", e);
        } catch (IllegalArgumentException e) {
            throw new WorkbookRecordException("cannot set String value on WorkbookRecord", e);
        }
    }

    private class ValueMethod {

        private final CellValueType valueType;

        private final String valueMethodName;

        private CellValueType autoValueType;

        private Method valueMethod;

        ValueMethod(@NotNull CellValueType valueType,
                    @NotNull String valueMethodName) {

            this.valueType = valueType;
            this.valueMethodName = valueMethodName;
        }

        @Nullable
        Object getValue(@Nullable E record, @NotNull CellValue cellValue) {
            if (column < 0) {
                throw new WorkbookRecordException("cannot get value for column " + column);
            }

            if (record == null) {
                return null;
            }

            final Method valueMethod = getValueMethod();
            if (valueMethod != null) {
                return getValueByMethod(valueMethod, record, cellValue);
            }

            return getValueByType(cellValue);
        }

        @Nullable
        private Object getValueByType(@NotNull CellValue cellValue) {
            try {
                return getValueByType0(cellValue);
            } catch (RuntimeException e) {
                if (e instanceof WorkbookRecordException) {
                    throw e;
                } else {
                    String msg = "cannot get value for field '" + field.getName() +
                            "' using CellValueType '" + valueType + "'";
                    throw new WorkbookRecordException(msg, e);
                }
            }
        }

        @Nullable
        private Object getValueByType0(@NotNull CellValue cellValue) {
            final Class<?> fieldType = field.getType();

            if (cellValue.isNull()) {
                if (!strict && fieldType.isPrimitive()) {
                    return HandlerUtils.getDefaultValue(fieldType);
                } else {
                    return null;
                }
            }

            final CellValueType valueType = getValueType();
            assert valueType != CellValueType.AUTO;

            final CellValue myCellValue = strict ? cellValue.strict() : cellValue;
            final Object ret;
            switch (valueType) {
                case BOOLEAN:
                    ret = myCellValue.booleanValue();
                    break;
                case INTEGER:
                    ret = myCellValue.intValue();
                    break;
                case LONG:
                    ret = myCellValue.longValue();
                    break;
                case DOUBLE:
                    ret = myCellValue.doubleValue();
                    break;
                case DECIMAL:
                    ret = myCellValue.bigDecimalValue();
                    break;
                case STRING:
                    ret = myCellValue.trim().stringValue();
                    break;
                case RAW_STRING:
                    ret = myCellValue.stringValue();
                    break;
                case TIME:
                    ret = myCellValue.localTimeValue();
                    break;
                case DATE:
                    ret = myCellValue.localDateValue();
                    break;
                case DATETIME:
                    ret = myCellValue.localDateTimeValue();
                    break;
                default:
                    throw new AssertionError("Shouldn't happen");
            }

            if (fieldType == ret.getClass()) {
                return ret;
            }

            if (!strict) {
                if (fieldType == boolean.class || fieldType == Boolean.class) {
                    return HandlerUtils.asBoolean(ret);
                } else if (fieldType == byte.class || fieldType == Byte.class) {
                    return HandlerUtils.asByte(ret);
                } else if (fieldType == short.class || fieldType == Short.class) {
                    return HandlerUtils.asShort(ret);
                } else if (fieldType == float.class || fieldType == Float.class) {
                    return HandlerUtils.asFloat(ret);
                } else if (fieldType == Date.class) {
                    return HandlerUtils.asDate(ret);
                } else if (fieldType == Time.class) {
                    return HandlerUtils.asSqlTime(ret);
                } else if (fieldType == Timestamp.class) {
                    return HandlerUtils.asSqlTimestamp(ret);
                } else if (fieldType == java.sql.Date.class) {
                    return HandlerUtils.asSqlDate(ret);
                } else if (fieldType == BigInteger.class) {
                    return HandlerUtils.asBigInteger(ret);
                }
            }

            throw new WorkbookRecordException("cannot get value for field '" + field.getName() +
                    "' using CellValueType '" + valueType + "'");
        }

        @NotNull
        private CellValueType getValueType() {
            if (CellValueType.AUTO != valueType) {
                return valueType;
            }
            if (autoValueType != null) {
                return autoValueType;
            }

            final Class<?> type = field.getType();
            if (type == boolean.class || type == Boolean.class) {
                autoValueType = CellValueType.BOOLEAN;
            } else if (type == int.class || type == Integer.class) {
                autoValueType = CellValueType.INTEGER;
            } else if (type == long.class || type == Long.class) {
                autoValueType = CellValueType.LONG;
            } else if (type == double.class || type == Double.class) {
                autoValueType = CellValueType.DOUBLE;
            } else if (type == BigDecimal.class) {
                autoValueType = CellValueType.DECIMAL;
            } else if (type == String.class) {
                autoValueType = CellValueType.STRING;
            } else if (type == LocalDate.class) {
                autoValueType = CellValueType.DATE;
            } else if (type == LocalTime.class) {
                autoValueType = CellValueType.TIME;
            } else if (type == LocalDateTime.class) {
                autoValueType = CellValueType.DATETIME;
            }

            if (!strict) {
                if (type == byte.class || type == Byte.class) {
                    autoValueType = CellValueType.INTEGER;
                } else if (type == short.class || type == Short.class) {
                    autoValueType = CellValueType.INTEGER;
                } else if (type == float.class || type == Float.class) {
                    autoValueType = CellValueType.DOUBLE;
                } else if (type == Date.class) {
                    autoValueType = CellValueType.DATETIME;
                } else if (type == Time.class) {
                    autoValueType = CellValueType.TIME;
                } else if (type == Timestamp.class) {
                    autoValueType = CellValueType.DATETIME;
                } else if (type == java.sql.Date.class) {
                    autoValueType = CellValueType.DATE;
                } else if (type == BigInteger.class) {
                    autoValueType = CellValueType.DECIMAL;
                }
            }

            if (autoValueType != null) {
                return autoValueType;
            }

            throw new WorkbookRecordException("cannot auto-determine CellValueType of field '" + field.getName() + "'");
        }

        @Nullable
        Object getValueByMethod(@NotNull Method valueMethod,
                                @NotNull E record, @NotNull CellValue cellValue) {

            try {
                return valueMethod.invoke(record, cellValue);
            } catch (IllegalAccessException e) {
                String msg = "valueMethod specified by @Workbook.Property '" + valueMethodName + "' cannot be accessed";
                throw new WorkbookRecordException(msg, e);
            } catch (InvocationTargetException e) {
                String msg = "valueMethod specified by @Workbook.Property '" + valueMethodName +
                        "' encountered an error";
                throw new WorkbookRecordException(msg, e.getTargetException());
            }
        }

        private Method getValueMethod() {
            if (valueMethodName.isEmpty()) {
                return null;
            }

            if (valueMethod == null) {
                final Method method;
                try {
                    method = recordClass.getMethod(valueMethodName, CellValue.class);
                } catch (NoSuchMethodException e) {
                    String msg = "valueMethod specified by @Workbook.Property '" + valueMethodName + "' is not found";
                    throw new WorkbookRecordException(msg, e);
                }
                valueMethod = method;
            }

            return valueMethod;
        }

    }

}
