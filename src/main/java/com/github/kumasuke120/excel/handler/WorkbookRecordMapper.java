package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.lang.reflect.Field;
import java.util.*;

/**
 * A mapper for mapping a workbook record to its properties.
 * <p>
 * This annotation is currently experimental and may be changed or removed in the future.
 * </p>
 */
@SuppressWarnings("BooleanMethodIsAlwaysInverted")
@ApiStatus.Experimental
class WorkbookRecordMapper<E> {

    private final Class<E> recordClass;

    private final WorkbookRecord recordAnnotation;

    private final Map<Integer, WorkbookRecordProperty<E>> properties;

    /**
     * Constructs a new {@code WorkbookRecordMapper} for the specified record class.
     *
     * @param recordClass the record class
     * @throws WorkbookRecordException if the specified class is not annotated with {@code @WorkbookRecord}
     */
    WorkbookRecordMapper(@NotNull(exception = NullPointerException.class) Class<E> recordClass) {
        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.recordAnnotation = findRecordAnnotation();
        this.properties = initProperties();
    }

    /**
     * Checks if the specified sheet index is beyond the range of the record.
     *
     * @param sheetIndex the sheet index to check
     * @return {@code true} if the specified sheet index is beyond the range of the record, {@code false} otherwise
     */
    boolean beyondRange(int sheetIndex) {
        return sheetIndex >= recordAnnotation.endSheet();
    }

    /**
     * Checks if the specified sheet index is within the range of the record.
     *
     * @param sheetIndex the sheet index to check
     * @return {@code true} if the specified sheet index is within the range of the record, {@code false} otherwise
     */
    boolean withinRange(int sheetIndex) {
        return sheetIndex >= recordAnnotation.startSheet() && sheetIndex < recordAnnotation.endSheet();
    }

    /**
     * Checks if the specified sheet index and row number are within the range of the record.
     *
     * @param sheetIndex the sheet index
     * @param rowNum     the row number
     * @return {@code true} if the specified sheet index and row number are within the range of the record,
     */
    boolean withinRange(int sheetIndex, int rowNum) {
        return withinRange(sheetIndex) &&
                rowNum >= recordAnnotation.startRow() && rowNum < recordAnnotation.endRow();
    }

    /**
     * Checks if the specified sheet index, row number, and column number are within the range of the record.
     *
     * @param sheetIndex the sheet index
     * @param rowNum     the row number
     * @param colNumNum  the column number
     * @return {@code true} if the specified sheet index, row number, and column number are within the range of the record,
     */
    boolean withinRange(int sheetIndex, int rowNum, int colNumNum) {
        return withinRange(sheetIndex, rowNum) &&
                colNumNum >= recordAnnotation.startColumn() && colNumNum < recordAnnotation.endColumn();
    }

    /**
     * Sets the value of the specified cell to the specified record based on the column number.
     *
     * @param record    the record to set
     * @param columnNum the column number of the cell
     * @param cellValue the value of the cell
     */
    void setValue(@NotNull E record, int columnNum, @NotNull CellValue cellValue) {
        final WorkbookRecordProperty<E> property = properties.get(columnNum);
        if (property == null) {
            return;
        }
        property.set(record, cellValue);
    }

    /**
     * Sets the sheet index of the specified record.
     *
     * @param record     the record to set
     * @param sheetIndex the sheet index to set
     */
    void setSheetIndex(@NotNull E record, int sheetIndex) {
        final WorkbookRecordProperty<E> property = properties.get(WorkbookRecordProperty.COLUMN_NUM_SHEET_INDEX);
        if (property == null) {
            return;
        }
        property.set(record, sheetIndex);
    }

    /**
     * Sets the sheet name of the specified record.
     *
     * @param record    the record to set
     * @param sheetName the sheet name to set
     */
    void setSheetName(@NotNull E record, @NotNull String sheetName) {
        final WorkbookRecordProperty<E> property = properties.get(WorkbookRecordProperty.COLUMN_NUM_SHEET_NAME);
        if (property == null) {
            return;
        }
        property.set(record, sheetName);
    }

    /**
     * Sets the row number of the specified record.
     *
     * @param record    the record to set
     * @param rowNumber the row number to set
     */
    void setRowNumber(@NotNull E record, int rowNumber) {
        final WorkbookRecordProperty<E> property = properties.get(WorkbookRecordProperty.COLUMN_NUM_ROW_NUMBER);
        if (property == null) {
            return;
        }
        property.set(record, rowNumber);
    }

    private WorkbookRecord findRecordAnnotation() {
        assert recordClass != null;
        WorkbookRecord annotation = recordClass.getAnnotation(WorkbookRecord.class);
        assert annotation != null;
        return annotation;
    }

    private Map<Integer, WorkbookRecordProperty<E>> initProperties() {
        final Field[] fields = recordClass.getDeclaredFields();

        final List<WorkbookRecordProperty<E>> properties = new ArrayList<>(fields.length);
        for (Field field : fields) {
            final FieldPropertyKind propertyKind = detectFieldPropertyKind(field);
            final WorkbookRecordProperty<E> prop;
            switch (propertyKind) {
                case METADATA:
                    prop = WorkbookRecordProperty.newMetadataProperty(recordClass, field);
                    break;
                case PROPERTY:
                    prop = WorkbookRecordProperty.newNormalProperty(recordClass, field);
                    break;
                default:
                    continue;
            }
            properties.add(prop);
        }

        return asPropertyMap(properties);
    }

    @NotNull
    private Map<Integer, WorkbookRecordProperty<E>> asPropertyMap(List<WorkbookRecordProperty<E>> properties) {

        final Map<Integer, WorkbookRecordProperty<E>> ret = new HashMap<>(properties.size());
        for (WorkbookRecordProperty<E> property : properties) {
            final boolean exists = ret.putIfAbsent(property.getColumn(), property) != null;
            if (exists) {
                if (property.getColumn() == WorkbookRecordProperty.COLUMN_NUM_SHEET_INDEX) {
                    throw new WorkbookRecordException("multiple @WorkbookRecord.Metadata(SHEET_INDEX) found");
                } else if (property.getColumn() == WorkbookRecordProperty.COLUMN_NUM_SHEET_NAME) {
                    throw new WorkbookRecordException("multiple @WorkbookRecord.Metadata(SHEET_NAME) found");
                } else if (property.getColumn() == WorkbookRecordProperty.COLUMN_NUM_ROW_NUMBER) {
                    throw new WorkbookRecordException("multiple @WorkbookRecord.Metadata(ROW_NUMBER) found");
                } else {
                    throw new WorkbookRecordException("column '" + property.getColumn() +
                            "' specified @WorkbookRecord.Property already exists");
                }
            }
        }
        return Collections.unmodifiableMap(ret);
    }

    @NotNull
    private FieldPropertyKind detectFieldPropertyKind(@NotNull Field field) {
        final WorkbookRecord.Metadata metaA = field.getAnnotation(WorkbookRecord.Metadata.class);
        final WorkbookRecord.Property propA = field.getAnnotation(WorkbookRecord.Property.class);
        if (propA != null && metaA != null) {
            throw new WorkbookRecordException("field cannot be annotated by both @WorkbookRecord.Metadata and " +
                    "@WorkbookRecord.Property");
        } else if (metaA != null) {
            return FieldPropertyKind.METADATA;
        } else if (propA != null) {
            return FieldPropertyKind.PROPERTY;
        } else {
            return FieldPropertyKind.NONE;
        }
    }

    private enum FieldPropertyKind {

        NONE,

        METADATA,

        PROPERTY

    }

}
