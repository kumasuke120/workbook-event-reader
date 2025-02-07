package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.util.StringUtils;
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

    private final PropertyBinder propertyBinder;

    /**
     * Constructs a new {@code WorkbookRecordMapper} for the specified record class.
     *
     * @param recordClass the record class
     * @throws WorkbookRecordException if the specified class is not annotated with {@code @WorkbookRecord}
     */
    WorkbookRecordMapper(@NotNull(exception = NullPointerException.class) Class<E> recordClass) {
        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.recordAnnotation = findRecordAnnotation();
        this.propertyBinder = initPropertyBinder();
    }

    private WorkbookRecord findRecordAnnotation() {
        return recordClass.getAnnotation(WorkbookRecord.class);
    }

    private PropertyBinder initPropertyBinder() {
        final Field[] fields = getDeclaredFields();

        final List<WorkbookRecordProperty<E>> properties = new ArrayList<>(fields.length);
        for (Field field : fields) {
            final WorkbookRecordProperty.Kind propertyKind = detectFieldPropertyKind(field);
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

        return checkAndInitPropertyBinder(properties);
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
     * Checks if the specified row number is the title row.
     *
     * @param rowNum the row number to check
     * @return {@code true} if the specified row number is the title row, {@code false} otherwise
     */
    boolean isTitleRow(int rowNum) {
        return rowNum >= 0 && rowNum == recordAnnotation.titleRow();
    }

    /**
     * Sets the value of the specified cell to the specified record based on the column number.
     *
     * @param record    the record to set
     * @param columnNum the column number of the cell
     * @param cellValue the value of the cell
     * @throws WorkbookRecordException if failed to set the value
     */
    void setValue(@NotNull E record, int columnNum, @NotNull CellValue cellValue) {
        final WorkbookRecordProperty<E> property = propertyBinder.getByColumn(columnNum);
        if (property == null) {
            return;
        }
        property.set(record, cellValue);
    }

    /**
     * Sets the value of the specified cell to the specified record based on the metadata type.
     *
     * @param record       the record to set
     * @param metadataType the metadata type
     * @param value        the value of the cell
     */
    void setMetadata(@NotNull E record, @NotNull WorkbookRecord.MetadataType metadataType, @NotNull Object value) {
        final WorkbookRecordProperty<E> property = propertyBinder.getByMetadata(metadataType);
        if (property == null) {
            return;
        }
        final WorkbookRecord.CellValueType valueType = metadataType.getValueType();
        switch (valueType) {
            case STRING:
                property.set(record, (String) value);
                break;
            case INTEGER:
                property.set(record, (int) value);
                break;
            default:
                throw new AssertionError("Shouldn't happen");
        }
    }

    /**
     * Gets all preset column titles of the record.
     *
     * @return all preset column titles of the record
     */
    TreeMap<Integer, String> getPresetColumnTitles() {
        return propertyBinder.getPropertyColumnTitles();
    }

    @NotNull
    private Field[] getDeclaredFields() {
        List<Field> fields = new ArrayList<>();
        Class<?> clazz;
        for (clazz = recordClass; clazz != null; clazz = clazz.getSuperclass()) {
            fields.addAll(Arrays.asList(clazz.getDeclaredFields()));
        }
        return fields.toArray(new Field[0]);
    }

    @NotNull
    private PropertyBinder checkAndInitPropertyBinder(List<WorkbookRecordProperty<E>> properties) {

        final Map<Integer, WorkbookRecordProperty<E>> map = new HashMap<>(properties.size());
        for (WorkbookRecordProperty<E> property : properties) {
            final boolean exists = map.putIfAbsent(property.getColumn(), property) != null;
            if (exists) {
                if (property.isMetadataType(WorkbookRecord.MetadataType.SHEET_INDEX)) {
                    throw new WorkbookRecordException("multiple @WorkbookRecord.Metadata(SHEET_INDEX) found");
                } else if (property.isMetadataType(WorkbookRecord.MetadataType.SHEET_NAME)) {
                    throw new WorkbookRecordException("multiple @WorkbookRecord.Metadata(SHEET_NAME) found");
                } else if (property.isMetadataType(WorkbookRecord.MetadataType.ROW_NUMBER)) {
                    throw new WorkbookRecordException("multiple @WorkbookRecord.Metadata(ROW_NUMBER) found");
                } else {
                    throw new WorkbookRecordException("A @WorkbookRecord.Property is already specified for column '" +
                            property.getColumn() + "'");
                }
            }
        }
        return new PropertyBinder(map);
    }

    @NotNull
    private WorkbookRecordProperty.Kind detectFieldPropertyKind(@NotNull Field field) {
        final WorkbookRecord.Metadata metaA = field.getAnnotation(WorkbookRecord.Metadata.class);
        final WorkbookRecord.Property propA = field.getAnnotation(WorkbookRecord.Property.class);
        if (propA != null && metaA != null) {
            throw new WorkbookRecordException("field cannot be annotated by both @WorkbookRecord.Metadata and " +
                    "@WorkbookRecord.Property");
        } else if (metaA != null) {
            return WorkbookRecordProperty.Kind.METADATA;
        } else if (propA != null) {
            return WorkbookRecordProperty.Kind.PROPERTY;
        } else {
            return WorkbookRecordProperty.Kind.NONE;
        }
    }

    /**
     * A binder for binding properties to columns.
     */
    final class PropertyBinder {

        private final Map<Integer, WorkbookRecordProperty<E>> properties;

        PropertyBinder(Map<Integer, WorkbookRecordProperty<E>> properties) {
            this.properties = Collections.unmodifiableMap(properties);
        }

        /**
         * Gets the property of the specified column.
         *
         * @param column the column number
         * @return the property of the specified column
         */
        WorkbookRecordProperty<E> getByColumn(int column) {
            return properties.get(column);
        }

        /**
         * Gets the property of the specified metadata type.
         *
         * @param metadataType the metadata type
         * @return the property of the specified metadata type
         */
        WorkbookRecordProperty<E> getByMetadata(WorkbookRecord.MetadataType metadataType) {
            return properties.get(metadataType.getMetaColumn());
        }

        /**
         * Gets the column titles of the normal properties.
         *
         * @return the column titles of the properties
         */
        TreeMap<Integer, String> getPropertyColumnTitles() {
            final TreeMap<Integer, String> columnTitles = new TreeMap<>();
            for (WorkbookRecordProperty<E> property : properties.values()) {
                if (property.isProperty()) {
                    final String title = property.getColumnTitle();
                    if (StringUtils.isNotEmpty(title)) {
                        columnTitles.put(property.getColumn(), title);
                    }
                }
            }
            return columnTitles;
        }

    }

}
