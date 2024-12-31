package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Metadata;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Property;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

class WorkbookRecordPropertyTest {

    @Test
    void setObjectValue() {
        final TestRecord record = new TestRecord();
        final WorkbookRecordProperty<TestRecord> property = newTestProperty("value");
        assertEquals(1, property.getColumn());

        final CellValue cellValue = newCellValue("test");
        property.set(record, cellValue);
        assertEquals("test", record.value);

        // null value
        final CellValue cellValue2 = newCellValue(null);
        property.set(record, cellValue2);
        assertNull(record.value);
    }

    @Test
    void setIntValue() {
        final TestRecord record = new TestRecord();

        final WorkbookRecordProperty<TestRecord> property = newTestProperty("intValue");
        assertEquals(2, property.getColumn());

        property.set(record, 123);
        assertEquals(123, record.intValue);

        // null value
        final CellValue cellValue2 = newCellValue(null);
        property.set(record, cellValue2);
        assertEquals(0, record.intValue);

        final WorkbookRecordProperty<TestRecord> property2 = newTestProperty("integerValue");
        assertEquals(3, property2.getColumn());
        property2.set(record, 456);
        assertEquals(456, record.integerValue);
    }

    @Test
    void setMetadataValue() {
        final TestRecord record = new TestRecord();
        final WorkbookRecordProperty<TestRecord> property = newTestMetadataProperty("rowNumber");
        property.set(record, 1);
        assertEquals(1, record.rowNumber);
    }

    private static WorkbookRecordProperty<TestRecord> newTestProperty(String fieldName) {
        try {
            final Field field = TestRecord.class.getDeclaredField(fieldName);
            return WorkbookRecordProperty.newNormalProperty(TestRecord.class, field);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    @SuppressWarnings("SameParameterValue")
    private static WorkbookRecordProperty<TestRecord> newTestMetadataProperty(String fieldName) {
        try {
            final Field field = TestRecord.class.getDeclaredField(fieldName);
            return WorkbookRecordProperty.newMetadataProperty(TestRecord.class, field);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    private static CellValue newCellValue(Object value) {
        try {
            final Method newInstanceMethod = CellValue.class.getDeclaredMethod("newInstance", Object.class);
            newInstanceMethod.setAccessible(true);
            final CellValue cellValue = (CellValue) newInstanceMethod.invoke(null, value);
            return cellValue.strict();
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    @SuppressWarnings("unused")
    @WorkbookRecord
    public static class TestRecord {
        @Metadata(MetadataType.ROW_NUMBER)
        private int rowNumber;
        @Property(column = 1)
        private String value;
        @Property(column = 2)
        private int intValue;
        @Property(column = 3)
        private Integer integerValue;
    }
}