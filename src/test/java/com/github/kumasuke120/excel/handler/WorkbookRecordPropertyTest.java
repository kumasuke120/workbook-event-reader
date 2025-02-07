package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.handler.WorkbookRecord.CellValueType;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Metadata;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Property;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigInteger;
import java.sql.Time;
import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.*;

class WorkbookRecordPropertyTest {

    @Test
    void determineValueType() {
        final WorkbookRecordProperty<TestRecord> property = newTestProperty("booleanValue");
        assertEquals(20, property.getColumn());
        assertUsingValueType(CellValueType.BOOLEAN, property);

        final WorkbookRecordProperty<TestRecord> property2 = newTestProperty("booleanValue2");
        assertEquals(21, property2.getColumn());
        assertUsingValueType(CellValueType.BOOLEAN, property2);

        final WorkbookRecordProperty<TestRecord> property3 = newTestProperty("longValue");
        assertEquals(22, property3.getColumn());
        assertUsingValueType(CellValueType.LONG, property3);

        final WorkbookRecordProperty<TestRecord> property4 = newTestProperty("longValue2");
        assertEquals(23, property4.getColumn());
        assertUsingValueType(CellValueType.LONG, property4);

        final WorkbookRecordProperty<TestRecord> property5 = newTestProperty("doubleValue");
        assertEquals(24, property5.getColumn());
        assertUsingValueType(CellValueType.DOUBLE, property5);

        final WorkbookRecordProperty<TestRecord> property6 = newTestProperty("doubleValue2");
        assertEquals(25, property6.getColumn());
        assertUsingValueType(CellValueType.DOUBLE, property6);

        final WorkbookRecordProperty<TestRecord> property7 = newTestProperty("localTimeValue");
        assertEquals(26, property7.getColumn());
        assertUsingValueType(CellValueType.TIME, property7);

        final WorkbookRecordProperty<TestRecord> property8 = newTestProperty("localDateTimeValue");
        assertEquals(27, property8.getColumn());
        assertUsingValueType(CellValueType.DATETIME, property8);

        assertThrows(WorkbookRecordException.class, () -> newTestProperty("objectValue"));
    }

    void assertUsingValueType(CellValueType expected, WorkbookRecordProperty<?> property) {
        if (property == null) {
            fail("property is null");
        } else if (property.valueMethod == null) {
            fail("property.valueMethod is null");
        } else {
            assertEquals(expected, property.valueMethod.usingValueType);
        }
    }

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
        property3.set(record, cellValue2);
        assertThrows(WorkbookRecordException.class, () -> property3.set(record, cellValue3));

        final CellValue cellValue4 = newCellValue(123L);
        final WorkbookRecordProperty<TestRecord> property4 = newTestProperty("strictIntegerValue");
        assertThrows(WorkbookRecordException.class, () -> property4.set(record, cellValue4));
    }

    @SuppressWarnings({"rawtypes", "unchecked"})
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

        assertThrows(WorkbookRecordException.class, () -> newTestProperty("plainValue"));

        final WorkbookRecordProperty property3 = newTestProperty("integerValue");
        assertThrows(WorkbookRecordException.class, () -> property3.set(new Object(), cellValue2));
    }

    @SuppressWarnings({"rawtypes", "unchecked"})
    @Test
    void setMetadataValue() {
        final TestRecord record = new TestRecord();
        final WorkbookRecordProperty<TestRecord> property = newTestMetadataProperty("rowNumber");
        property.set(record, 1);
        assertEquals(1, record.rowNumber);

        final WorkbookRecordProperty<TestRecord> property2 = newTestMetadataProperty("sheetIndex");
        property2.set(record, 2);
        assertEquals(2, record.sheetIndex);

        final WorkbookRecordProperty property3 = newTestMetadataProperty("sheetIndex");
        assertThrows(WorkbookRecordException.class, () -> property3.set(new Object(), 2));

        final WorkbookRecordProperty<TestRecord> property4 = newTestMetadataProperty("sheetName");
        property4.set(record, "name");
        assertEquals("name", record.sheetName);

        final WorkbookRecordProperty property5 = newTestMetadataProperty("sheetName");
        assertThrows(WorkbookRecordException.class, () -> property5.set(new Object(), "name"));

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

        final WorkbookRecordProperty<TestRecord> property3 = newTestProperty("errorProneMethod");
        assertEquals(28, property3.getColumn());
        final CellValue cellValue2 = newCellValue("abc");
        assertThrows(WorkbookRecordException.class, () -> property3.set(record, cellValue2));
    }

    @Test
    void setLenientValue() {
        final TestRecord record = new TestRecord();
        final CellValue numberValue = newCellValue("1");

        final WorkbookRecordProperty<TestRecord> byteProperty = newTestProperty("lenientByteValue");
        byteProperty.set(record, numberValue);
        assertEquals(1, record.lenientByteValue);

        final WorkbookRecordProperty<TestRecord> byteProperty2 = newTestProperty("lenientByteValue2");
        byteProperty2.set(record, numberValue);
        assertEquals(Byte.valueOf((byte) 1), record.lenientByteValue2);

        final WorkbookRecordProperty<TestRecord> shortProperty = newTestProperty("lenientShortValue");
        shortProperty.set(record, numberValue);
        assertEquals(1, record.lenientShortValue);

        final WorkbookRecordProperty<TestRecord> shortProperty2 = newTestProperty("lenientShortValue2");
        shortProperty2.set(record, numberValue);
        assertEquals(Short.valueOf((short) 1), record.lenientShortValue2);

        final WorkbookRecordProperty<TestRecord> floatProperty = newTestProperty("lenientFloatValue");
        floatProperty.set(record, numberValue);
        assertEquals(1.0f, record.lenientFloatValue, 1e-6);

        final WorkbookRecordProperty<TestRecord> floatProperty2 = newTestProperty("lenientFloatValue2");
        floatProperty2.set(record, numberValue);
        assertEquals(1.0f, record.lenientFloatValue2, 1e-6);

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
        assertEquals(BigInteger.ONE, record.lenientBigIntegerValue);

        final WorkbookRecordProperty<TestRecord> integerObjectValue = newTestProperty("integerObjectValue");
        integerObjectValue.set(record, numberValue);
        assertEquals(1, record.integerObjectValue);
    }

    @Test
    void columnTitle() {
        final WorkbookRecordProperty<TestRecord> integerObjectValue = newTestProperty("integerObjectValue");
        assertEquals("", integerObjectValue.getColumnTitle());

        final WorkbookRecordProperty<TestRecord> titleValue = newTestProperty("titleValue");
        assertEquals("test title", titleValue.getColumnTitle());

        final WorkbookRecordProperty<TestRecord> sheetIndex = newTestMetadataProperty("sheetIndex");
        assertThrows(WorkbookRecordException.class, sheetIndex::getColumnTitle);
    }

    private static WorkbookRecordProperty<TestRecord> newTestProperty(String fieldName) {
        try {
            final Field field = TestRecord.class.getDeclaredField(fieldName);
            final WorkbookRecordProperty<TestRecord> property = WorkbookRecordProperty.newNormalProperty(TestRecord.class, field);
            assertDoesNotThrow(property::propertyAnnotation);
            assertThrows(WorkbookRecordException.class, property::metadataAnnotation);
            return property;
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    @SuppressWarnings("SameParameterValue")
    private static WorkbookRecordProperty<TestRecord> newTestMetadataProperty(String fieldName) {
        try {
            final Field field = TestRecord.class.getDeclaredField(fieldName);
            final WorkbookRecordProperty<TestRecord> property = WorkbookRecordProperty.newMetadataProperty(TestRecord.class, field);
            assertDoesNotThrow(property::metadataAnnotation);
            assertThrows(WorkbookRecordException.class, property::propertyAnnotation);
            return property;
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
        @SuppressWarnings({"InjectedReferences", "RedundantSuppression"})
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
        @Property(column = 16)
        private Byte lenientByteValue2;
        @Property(column = 17)
        private Short lenientShortValue2;
        @Property(column = 18)
        private Float lenientFloatValue2;
        @Property(column = 19)
        private Object objectValue;

        @Property(column = 20)
        private boolean booleanValue;
        @Property(column = 21)
        private Boolean booleanValue2;
        @Property(column = 22)
        private long longValue;
        @Property(column = 23)
        private Long longValue2;
        @Property(column = 24)
        private double doubleValue;
        @Property(column = 25)
        private Double doubleValue2;
        @Property(column = 26)
        private LocalTime localTimeValue;
        @Property(column = 27)
        private LocalDateTime localDateTimeValue;

        @Property(column = 28, valueMethod = "errorProneMethod")
        private Integer errorProneMethod;

        @Property(column = 29, valueType = CellValueType.INTEGER)
        private Object integerObjectValue;

        @Property(column = 30, valueType = CellValueType.LONG, strict = true)
        private Integer strictIntegerValue;

        @Property(column = 31, title = "test title")
        private String titleValue;

        private Object privateMethod(CellValue cellValue) {
            throw new UnsupportedOperationException();
        }

        public Object absIntValue(CellValue cellValue) {
            return Math.abs(cellValue.intValue());
        }

        public Integer errorProneMethod(CellValue cellValue) {
            return Integer.parseInt(cellValue.stringValue());
        }
    }
}