package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.lang.annotation.*;

/**
 * An annotation that marks a class as a workbook record.
 * <p>
 * This annotation is currently experimental and may be changed or removed in the future.
 * </p>
 */
@ApiStatus.Experimental
@Documented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface WorkbookRecord {

    /**
     * The index(zero-based, inclusive) of the first sheet to be processed.
     *
     * @return the index of the first sheet to be processed
     */
    int startSheet() default 0;

    /**
     * The index(zero-based, exclusive) of the last sheet to be processed.
     *
     * @return the index of the last sheet to be processed
     */
    int endSheet() default 255;

    /**
     * The index(zero-based, inclusive) of the first row to be processed.
     *
     * @return the index of the first row to be processed
     */
    int startRow() default 0;

    /**
     * The index(zero-based, exclusive) of the last row to be processed.
     *
     * @return the index of the last row to be processed
     */
    int endRow() default Integer.MAX_VALUE;

    /**
     * The index(zero-based, inclusive) of the first column to be processed.
     *
     * @return the index of the first column to be processed
     */
    int startColumn() default 0;

    /**
     * The index(zero-based, exclusive) of the last column to be processed.
     *
     * @return the index of the last column to be processed
     */
    int endColumn() default 16384;

    /**
     * An annotation that marks a field as a metadata field.
     */
    @Documented
    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    @interface Metadata {

        /**
         * The type of the metadata.
         *
         * @return the type of the metadata
         */
        MetadataType value();

    }

    /**
     * The type of the metadata.
     */
    enum MetadataType {

        /**
         * The index(zero-based) of the sheet.
         */
        SHEET_INDEX(-1, CellValueType.INTEGER),

        /**
         * The name of the sheet.
         */
        SHEET_NAME(-2, CellValueType.STRING),

        /**
         * The index(zero-based) of the row.
         */
        ROW_NUMBER(-3, CellValueType.INTEGER),

        ;

        private final int metaColumn;
        private final CellValueType valueType;

        MetadataType(int metaColumn, CellValueType valueType) {
            this.metaColumn = metaColumn;
            this.valueType = valueType;
        }

        /**
         * Gets the special column bound to the metadata type.
         *
         * @return the column bound to the metadata type
         */
        public int getMetaColumn() {
            return metaColumn;
        }

        /**
         * Gets the value type of the metadata value.
         *
         * @return the value type of the metadata value
         */
        public CellValueType getValueType() {
            return valueType;
        }

        <E> void setMetadataValue(WorkbookRecordProperty<E> property, E record, Object value) {
            if (valueType == CellValueType.STRING) {
                property.set(record, (String) value);
            } else if (valueType == CellValueType.INTEGER) {
                property.set(record, (int) value);
            } else {
                throw new AssertionError("Shouldn't happen");
            }
        }
    }

    /**
     * An annotation that marks a field as a property field.
     */
    @Documented
    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    @interface Property {

        /**
         * The index(zero-based) of the column bound to.
         *
         * @return the index of the column bound
         */
        int column();

        /**
         * Whether the value should be strictly checked when getting from the workbook.
         *
         * @return {@code true} if the value should be strictly checked, {@code false} otherwise
         * @see CellValue#strict()
         */
        boolean strict() default false;

        /**
         * Whether the value should be trimmed when getting from the workbook if it is a string.
         *
         * @return {@code true} if the value should be trimmed, {@code false} otherwise
         * @see CellValue#trim()
         */
        boolean trim() default true;

        /**
         * The type of the cell value.
         *
         * @return the type of the cell value
         */
        CellValueType valueType() default CellValueType.AUTO;

        /**
         * The method name to get the value from the cell value.
         * <p>
         * * If this value is empty, the getter method of the marked filed will be used.<br>
         * * If this value is not empty, the method(within the record class) with the specified name will be used.<br>
         * * The method must take a {@link CellValue} as the only parameter and returns the value of any type.<br>
         * * The method must be public and non-static.
         * </p>
         *
         * @return the method name to get the value from the cell value
         */
        String valueMethod() default "";

    }

    /**
     * The type of the cell value.
     */
    enum CellValueType {

        /**
         * Automatically detect the type of the cell value.
         */
        AUTO,

        /**
         * The cell value is a <code>boolean</code>.
         * <p>
         * In non-strict mode, any value that can be converted from a <code>boolean</code> is acceptable.<br>
         * In strict mode, only the value that is exactly a <code>boolean</code> is acceptable.
         * </p>
         */
        BOOLEAN {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.booleanValue();
            }
        },

        /**
         * The cell value is an <code>int</code>.
         * <p>
         * In non-strict mode, any value that can be converted from a <code>int</code> is acceptable.<br>
         * In strict mode, only the value that is exactly a <code>int</code> is acceptable.
         * </p>
         */
        INTEGER {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.intValue();
            }
        },

        /**
         * The cell value is a <code>long</code>.
         * <p>
         * In non-strict mode, any value that can be converted from a <code>long</code> is acceptable.<br>
         * In strict mode, only the value that is exactly a <code>long</code> is acceptable.
         * </p>
         */
        LONG {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.longValue();
            }
        },

        /**
         * The cell value is a <code>double</code>.
         * <p>
         * In non-strict mode, any value that can be converted from a <code>double</code> is acceptable.<br>
         * In strict mode, only the value that is exactly a <code>double</code> is acceptable.
         * </p>
         */
        DOUBLE {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.doubleValue();
            }
        },

        /**
         * The cell value is a {@link java.math.BigDecimal}.
         * <p>
         * In non-strict mode, any value that can be converted from a {@link java.math.BigDecimal} is acceptable.<br>
         * In strict mode, only the value that is exactly a {@link java.math.BigDecimal} is acceptable.
         * </p>
         */
        DECIMAL {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.bigDecimalValue();
            }
        },

        /**
         * The cell value is a {@link String}.
         * <p>
         * In non-strict mode, any value that can be converted from a {@link String} is acceptable.<br>
         * In strict mode, only the value that is exactly a {@link String} is acceptable.
         * </p>
         */
        STRING {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.stringValue();
            }
        },

        /**
         * The cell value is a {@link java.time.LocalTime}.
         * <p>
         * In non-strict mode, any value that can be converted from a {@link java.time.LocalTime} is acceptable.<br>
         * In strict mode, only the value that is exactly a {@link java.time.LocalTime} is acceptable.
         * </p>
         */
        TIME {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.localTimeValue();
            }
        },

        /**
         * The cell value is a {@link java.time.LocalDate}.
         * <p>
         * In non-strict mode, any value that can be converted from a {@link java.time.LocalDate} is acceptable.<br>
         * In strict mode, only the value that is exactly a {@link java.time.LocalDate} is acceptable.
         * </p>
         */
        DATE {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.localDateValue();
            }
        },

        /**
         * The cell value is a {@link java.time.LocalDateTime}.
         * <p>
         * In non-strict mode, any value that can be converted from a {@link java.time.LocalDateTime} is acceptable.<br>
         * In strict mode, only the value that is exactly a {@link java.time.LocalDateTime} is acceptable.
         * </p>
         */
        DATETIME {
            @Override
            public Object getValue(@NotNull CellValue cellValue) {
                return cellValue.localDateTimeValue();
            }
        },

        ;

        /**
         * Gets the value from the specified {@link CellValue} based on the current type.
         *
         * @param cellValue the cell value
         * @return the value
         */
        public Object getValue(@NotNull CellValue cellValue) {
            throw new UnsupportedOperationException();
        }

    }


}
