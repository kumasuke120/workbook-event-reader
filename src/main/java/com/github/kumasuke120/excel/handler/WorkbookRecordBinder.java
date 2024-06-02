package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.lang.reflect.Field;
import java.util.*;

@SuppressWarnings("BooleanMethodIsAlwaysInverted")
@ApiStatus.Experimental
class WorkbookRecordBinder<E> {

    private final Class<E> recordClass;

    private final WorkbookRecord recordAnnotation;

    private final Map<Integer, WorkbookRecordProperty<E>> properties;

    WorkbookRecordBinder(@NotNull(exception = NullPointerException.class) Class<E> recordClass) {
        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.recordAnnotation = findAnnotation();
        this.properties = initProperties();
    }

    boolean withinRange(int sheetIndex) {
        return sheetIndex >= recordAnnotation.startSheet() && sheetIndex < recordAnnotation.endSheet();
    }

    boolean withinRange(int sheetIndex, int rowNum) {
        return withinRange(sheetIndex) &&
                rowNum >= recordAnnotation.startRow() && rowNum < recordAnnotation.endRow();
    }

    boolean withinRange(int sheetIndex, int rowNum, int colNumNum) {
        return withinRange(sheetIndex, rowNum) &&
                colNumNum >= recordAnnotation.startColumn() && colNumNum < recordAnnotation.endColumn();
    }

    void setValue(@NotNull E record, int columnNum, @NotNull CellValue cellValue) {
        WorkbookRecordProperty<E> property = properties.get(columnNum);
        if (property == null) {
            return;
        }

        property.set(record, cellValue);
    }

    void setSheetIndex(@NotNull E record, int sheetIndex) {
        WorkbookRecordProperty<E> property = properties.get(WorkbookRecordProperty.COLUMN_NUM_SHEET_INDEX);
        if (property == null) {
            return;
        }

        property.set(record, sheetIndex);
    }

    void setRowNumber(@NotNull E record, int rowNumber) {
        WorkbookRecordProperty<E> property = properties.get(WorkbookRecordProperty.COLUMN_NUM_ROW_NUMBER);
        if (property == null) {
            return;
        }

        property.set(record, rowNumber);
    }

    private WorkbookRecord findAnnotation() {
        assert recordClass != null;
        WorkbookRecord annotation = recordClass.getAnnotation(WorkbookRecord.class);
        assert annotation != null;
        return annotation;
    }

    private Map<Integer, WorkbookRecordProperty<E>> initProperties() {
        final Field[] fields = recordClass.getDeclaredFields();

        final List<WorkbookRecordProperty<E>> properties = new ArrayList<>(fields.length);
        for (Field field : fields) {
            final WorkbookRecordProperty<E> normalProperty = initProperty(field);
            if (normalProperty != null) {
                properties.add(normalProperty);
            }

            final WorkbookRecordProperty<E> sheetIndexProperty = initSheetIndexProperty(field);
            if (sheetIndexProperty != null) {
                properties.add(sheetIndexProperty);
            }

            final WorkbookRecordProperty<E> rowNumberProperty = initRowNumberProperty(field);
            if (rowNumberProperty != null) {
                properties.add(rowNumberProperty);
            }
        }

        return asProperties(properties);
    }

    private @NotNull Map<Integer, WorkbookRecordProperty<E>>
    asProperties(List<WorkbookRecordProperty<E>> properties) {

        final Map<Integer, WorkbookRecordProperty<E>> ret = new HashMap<>(properties.size());
        for (WorkbookRecordProperty<E> property : properties) {
            boolean exists = ret.putIfAbsent(property.getColumn(), property) != null;
            if (exists) {
                if (property.getColumn() == WorkbookRecordProperty.COLUMN_NUM_SHEET_INDEX) {
                    throw new WorkbookRecordException("multiple @WorkbookRecord.SheetIndex found");
                } else if (property.getColumn() == WorkbookRecordProperty.COLUMN_NUM_ROW_NUMBER) {
                    throw new WorkbookRecordException("multiple @WorkbookRecord.RowNumber found");
                } else {
                    String msg = "column '" + property.getColumn() +
                            "' specified @WorkbookRecord.Property already exists";
                    throw new WorkbookRecordException(msg);
                }
            }
        }
        return Collections.unmodifiableMap(ret);
    }

    @Nullable
    private WorkbookRecordProperty<E> initSheetIndexProperty(@NotNull Field field) {
        WorkbookRecord.SheetIndex propA = field.getAnnotation(WorkbookRecord.SheetIndex.class);
        if (propA == null) {
            return null;
        }

        return new WorkbookRecordProperty<>(
                recordClass,
                WorkbookRecordProperty.COLUMN_NUM_SHEET_INDEX,
                "",
                true,
                WorkbookRecord.CellValueType.INTEGER,
                "",
                field
        );
    }

    @Nullable
    private WorkbookRecordProperty<E> initRowNumberProperty(@NotNull Field field) {
        WorkbookRecord.RowNumber propA = field.getAnnotation(WorkbookRecord.RowNumber.class);
        if (propA == null) {
            return null;
        }

        return new WorkbookRecordProperty<>(
                recordClass,
                WorkbookRecordProperty.COLUMN_NUM_ROW_NUMBER,
                "",
                true,
                WorkbookRecord.CellValueType.INTEGER,
                "",
                field
        );
    }

    @Nullable
    private WorkbookRecordProperty<E> initProperty(@NotNull Field field) {
        WorkbookRecord.Property propA = field.getAnnotation(WorkbookRecord.Property.class);
        if (propA == null) {
            return null;
        }

        return new WorkbookRecordProperty<>(
                recordClass,
                propA.column(),
                propA.title() == null ? "" : propA.title(),
                propA.strict(),
                propA.valueType() == null ? WorkbookRecord.CellValueType.AUTO : propA.valueType(),
                propA.valueMethod() == null ? "" : propA.valueMethod(),
                field
        );
    }

}
