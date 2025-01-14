package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Metadata;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Property;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigInteger;
import java.sql.Time;
import java.sql.Timestamp;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.*;

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

        // strict value
        final CellValue cellValue3 = newCellValue(" 123 ");
        final WorkbookRecordProperty<TestRecord> property3 = newTestProperty("strictLongValue");
        assertEquals(7, property3.getColumn());

        assertThrows(WorkbookRecordException.class, () -> property3.set(record, cellValue3));
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

        final WorkbookRecordProperty<TestRecord> property2 = newTestMetadataProperty("sheetIndex");
        property2.set(record, 2);
        assertEquals(2, record.sheetIndex);

        final WorkbookRecordProperty<TestRecord> property3 = newTestMetadataProperty("sheetName");
        property3.set(record, "name");
        assertEquals("name", record.sheetName);

        assertThrows(WorkbookRecordException.class, () -> newTestMetadataProperty("plainValue"));
    }

    @Test
    void setByValueMethod() {
        final TestRecord record = new TestRecord();
        assertThrows(WorkbookRecordException.class, () -> newTestProperty("notFoundMethod"));

        final CellValue cellValue = newCellValue(-123);
        final WorkbookRecordProperty<TestRecord> property = newTestProperty("privateMethod");
        assertEquals(5, property.getColumn());
        assertThrows(WorkbookRecordException.class, () -> property.set(record, cellValue));

        final WorkbookRecordProperty<TestRecord> property2 = newTestProperty("absIntValue");
        assertEquals(6, property2.getColumn());

        property2.set(record, cellValue);
        assertEquals(123, record.absIntValue);
    }

    @Test
    void setLenientValue() {
        final TestRecord record = new TestRecord();
        final CellValue numberValue = newCellValue("1");

        final WorkbookRecordProperty<TestRecord> byteProperty = newTestProperty("lenientByteValue");
        byteProperty.set(record, numberValue);

        final WorkbookRecordProperty<TestRecord> shortProperty = newTestProperty("lenientShortValue");
        shortProperty.set(record, numberValue);

        final WorkbookRecordProperty<TestRecord> floatProperty = newTestProperty("lenientFloatValue");
        floatProperty.set(record, numberValue);

        final CellValue dateTimeValue = newCellValue("2020-01-01 01:23:45");
        final WorkbookRecordProperty<TestRecord> dateProperty = newTestProperty("lenientDateValue");
        dateProperty.set(record, dateTimeValue);

        final WorkbookRecordProperty<TestRecord> timeProperty = newTestProperty("lenientTimeValue");
        timeProperty.set(record, dateTimeValue);

        final WorkbookRecordProperty<TestRecord> timestampProperty = newTestProperty("lenientTimestampValue");
        timestampProperty.set(record, dateTimeValue);

        final WorkbookRecordProperty<TestRecord> sqlDateProperty = newTestProperty("lenientSqlDateValue");
        sqlDateProperty.set(record, dateTimeValue);

        final WorkbookRecordProperty<TestRecord> bigIntegerProperty = newTestProperty("lenientBigIntegerValue");
        bigIntegerProperty.set(record, numberValue);
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
        @Metadata(MetadataType.SHEET_INDEX)
        private Integer sheetIndex;
        @Metadata(MetadataType.SHEET_NAME)
        private String sheetName;
        @Metadata(MetadataType.ROW_NUMBER)
        private int rowNumber;
        @Property(column = 1)
        private String value;
        @Property(column = 2)
        private int intValue;
        @Property(column = 3)
        private Integer integerValue;
        @SuppressWarnings("InjectedReferences")
        @Property(column = 4, valueMethod = "notFoundMethod")
        private Integer notFoundMethod;
        @Property(column = 5, valueMethod = "privateMethod")
        private Integer privateMethod;
        @Property(column = 6, valueMethod = "absIntValue")
        private Integer absIntValue;
        private Integer plainValue;
        @Property(column = 7, strict = true, trim = false)
        private Long strictLongValue;

        @Property(column = 8)
        private byte lenientByteValue;
        @Property(column = 9)
        private short lenientShortValue;
        @Property(column = 10)
        private float lenientFloatValue;
        @Property(column = 11)
        private Date lenientDateValue;
        @Property(column = 12)
        private Time lenientTimeValue;
        @Property(column = 13)
        private Timestamp lenientTimestampValue;
        @Property(column = 14)
        private java.sql.Date lenientSqlDateValue;
        @Property(column = 15)
        private BigInteger lenientBigIntegerValue;

        private Object privateMethod(CellValue cellValue) {
            throw new UnsupportedOperationException();
        }

        public Object absIntValue(CellValue cellValue) {
            return Math.abs(cellValue.intValue());
        }
    }
}