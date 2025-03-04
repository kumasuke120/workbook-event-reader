package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.handler.WorkbookRecord.CellValueType;
import com.github.kumasuke120.excel.util.StringUtils;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.lang.annotation.Annotation;
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
import java.util.Date;

/**
 * A class represents a property marked by {@link WorkbookRecord.Property} in a {@link WorkbookRecord}-marked class.
 * <p>
 * This class is intended for internal use only.
 * </p>
 */
@ApiStatus.Internal
final class WorkbookRecordProperty<E> {

    private final Class<E> recordClass;

    private final Field field;

    private final Kind kind;

    private final Annotation annotation;

    final ValueMethod valueMethod;

    WorkbookRecordProperty(@NotNull Class<E> recordClass, @NotNull Field field,
                           @NotNull WorkbookRecord.Property annotation) {
        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.field = field;
        this.kind = Kind.PROPERTY;
        this.annotation = annotation;
        this.valueMethod = new ValueMethod(annotation.valueType(), annotation.valueMethod());
    }

    WorkbookRecordProperty(@NotNull Class<E> recordClass, @NotNull Field field,
                           @NotNull WorkbookRecord.Metadata annotation) {
        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.field = field;
        this.kind = Kind.METADATA;
        this.annotation = annotation;
        this.valueMethod = new ValueMethod(annotation.value().getValueType(), "");
    }

    /**
     * Creates a new instance of {@code WorkbookRecordProperty} for metadata properties.
     *
     * @param recordClass the class of the record
     * @param field       the field of the metadata
     * @param <T>         the type of the record
     * @return a new instance of {@code WorkbookRecordProperty}
     */
    static <T> WorkbookRecordProperty<T> newMetadataProperty(@NotNull Class<T> recordClass, @NotNull Field field) {
        final WorkbookRecord.Metadata metaA = field.getAnnotation(WorkbookRecord.Metadata.class);
        if (metaA == null) {
            throw new WorkbookRecordException("field must be annotated with @WorkbookRecord.Metadata");
        }
        return new WorkbookRecordProperty<>(recordClass, field, metaA);
    }

    /**
     * Creates a new instance of {@code WorkbookRecordProperty} for normal properties.
     *
     * @param recordClass the class of the record
     * @param field       the field of the property
     * @param <T>         the type of the record
     * @return a new instance of {@code WorkbookRecordProperty}
     */
    static <T> WorkbookRecordProperty<T> newNormalProperty(@NotNull Class<T> recordClass, @NotNull Field field) {
        final WorkbookRecord.Property propA = field.getAnnotation(WorkbookRecord.Property.class);
        if (propA == null) {
            throw new WorkbookRecordException("field must be annotated with @WorkbookRecord.Property");
        }
        return new WorkbookRecordProperty<>(recordClass, field, propA);
    }

    /**
     * Converts the annotation to a property annotation if possible.
     *
     * @return the property annotation
     */
    WorkbookRecord.Property propertyAnnotation() {
        if (kind != Kind.PROPERTY) {
            throw new WorkbookRecordException("property is not a normal property");
        }
        return (WorkbookRecord.Property) annotation;
    }

    /**
     * Converts the annotation to a metadata annotation if possible.
     *
     * @return the metadata annotation
     */
    WorkbookRecord.Metadata metadataAnnotation() {
        if (kind != Kind.METADATA) {
            throw new WorkbookRecordException("property is not a metadata property");
        }
        return (WorkbookRecord.Metadata) annotation;
    }

    /**
     * Checks whether the property is in strict mode.
     *
     * @return {@code true} if the property is in strict mode, otherwise {@code false}
     */
    boolean isStrict() {
        return propertyAnnotation().strict();
    }

    /**
     * Checks whether the property is in trim mode.
     *
     * @return {@code true} if the property is in trim mode, otherwise {@code false}
     */
    boolean isTrim() {
        return propertyAnnotation().trim();
    }

    /**
     * Gets the column number of the property.
     *
     * @return the column number of the property
     */
    int getColumn() {
        switch (kind) {
            case PROPERTY:
                return propertyAnnotation().column();
            case METADATA:
                return metadataAnnotation().value().getMetaColumn();
            default:
                throw new AssertionError("Shouldn't happen");
        }
    }

    /**
     * Determines whether the property is a metadata property.
     *
     * @param metadataType the metadata type to check
     * @return {@code true} if the property is a metadata property, otherwise {@code false}
     */
    boolean isMetadataType(@NotNull WorkbookRecord.MetadataType metadataType) {
        if (kind != Kind.METADATA) {
            return false;
        }
        return metadataType == metadataAnnotation().value();
    }

    /**
     * Sets the value of the property from the specified record.
     *
     * @param record    the record to set the value to
     * @param cellValue the cell value to get the value from
     */
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

    /**
     * Sets the value of the property from the specified record.
     *
     * @param record the record to set the value to
     * @param value  the value to set
     */
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

    /**
     * Sets the value of the property from the specified record.
     *
     * @param record the record to set the value to
     * @param value  the value to set
     */
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

    /**
     * A class represents the method to get the value of the property.
     */
    class ValueMethod {

        final CellValueType valueType;

        final CellValueType usingValueType;

        final String valueMethodName;

        final Method valueMethod;

        /**
         * Creates a new instance of {@code ValueMethod}.
         *
         * @param valueType       the type of the value
         * @param valueMethodName the method name to get the value if specified
         * @throws WorkbookRecordException if value type cannot be determined or if the value method cannot be found
         */
        ValueMethod(@NotNull CellValueType valueType, @NotNull String valueMethodName) {
            this.valueType = valueType;
            this.valueMethodName = valueMethodName;
            this.usingValueType = determineValueType();
            this.valueMethod = findValueMethod();
        }

        @NotNull
        private CellValueType determineValueType() {
            if (CellValueType.AUTO != valueType) {
                return valueType;
            }
            if (StringUtils.isNotBlank(valueMethodName)) {
                return CellValueType.AUTO;
            }

            CellValueType autoValueType = null;
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

            if (!isStrict()) {
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

        private Method findValueMethod() {
            if (StringUtils.isBlank(valueMethodName)) {
                return null;
            }

            try {
                return recordClass.getDeclaredMethod(valueMethodName, CellValue.class);
            } catch (NoSuchMethodException e) {
                String msg = "valueMethod specified by @WorkbookRecord.Property cannot be found: " +
                        "Object " + valueMethodName + "(CellValue cellValue)";
                throw new WorkbookRecordException(msg, e);
            }
        }

        /**
         * Gets the value of the property from the specified {@link CellValue}.
         *
         * @param record    the given record, get the value method if necessary
         * @param cellValue the cell value to get the value from
         * @return the value of the property
         */
        @Nullable
        Object getValue(@NotNull E record, @NotNull CellValue cellValue) {
            if (valueMethod == null) {
                return getValueByType(cellValue);
            } else {
                return getValueByMethod(record, cellValue);
            }
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
                            "' using CellValueType '" + usingValueType + "'";
                    throw new WorkbookRecordException(msg, e);
                }
            }
        }

        @Nullable
        private Object getValueByType0(@NotNull CellValue cellValue) {
            final Class<?> fieldType = field.getType();

            if (cellValue.isNull()) {
                if (!isStrict() && fieldType.isPrimitive()) {
                    return HandlerUtils.getDefaultValue(fieldType);
                } else {
                    return null;
                }
            }

            assert usingValueType != CellValueType.AUTO;

            final CellValue preparedValue = prepareCellValue(cellValue);
            final Object ret = usingValueType.getValue(preparedValue);
            // if the field type matches the default type, return it directly
            if (HandlerUtils.primitiveToWrapper(fieldType) == ret.getClass()) {
                return ret;
            }

            if (!isStrict()) {
                if (fieldType == byte.class || fieldType == Byte.class) {
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
                } else if (fieldType == Object.class) {
                    return ret;
                }
            }

            throw new WorkbookRecordException("cannot get value for field '" + field.getName() +
                    "' using CellValueType '" + usingValueType + "'");
        }

        @NotNull
        private CellValue prepareCellValue(@NotNull CellValue cellValue) {
            CellValue ret = cellValue;
            if (isStrict()) {
                ret = ret.strict();
            }
            if (isTrim()) {
                ret = ret.trim();
            }
            return ret;
        }

        @Nullable
        private Object getValueByMethod(@NotNull E record, @NotNull CellValue cellValue) {
            try {
                final CellValue preparedValue = prepareCellValue(cellValue);
                return valueMethod.invoke(record, preparedValue);
            } catch (IllegalAccessException e) {
                String msg = "valueMethod specified by @WorkbookRecord.Property '" + valueMethodName + "' " +
                        "should be public: public Object " + valueMethodName + "(CellValue cellValue)";
                throw new WorkbookRecordException(msg, e);
            } catch (InvocationTargetException e) {
                String msg = "valueMethod specified by @WorkbookRecord.Property '" + valueMethodName +
                        "' encountered an error";
                throw new WorkbookRecordException(msg, e.getTargetException());
            }
        }

    }


    /**
     * An enumeration represents the kind of the property.
     */
    public enum Kind {

        /**
         * Represents an ignored property
         */
        NONE,

        /**
         * Represents a metadata property
         */
        METADATA,

        /**
         * Represents a normal property
         */
        PROPERTY

    }

}
