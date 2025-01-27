package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Property;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

@SuppressWarnings({"DataFlowIssue", "unused"})
class WorkbookRecordMapperTest {

    @Test
    void create() {
        assertThrows(NullPointerException.class, () -> new WorkbookRecordMapper<>(null));
        assertThrows(WorkbookRecordException.class, () -> new WorkbookRecordMapper<>(ErrorTestRecord.class));
        assertDoesNotThrow(() -> new WorkbookRecordMapper<>(TestRecord.class));

        assertThrows(WorkbookRecordException.class, () -> new WorkbookRecordMapper<>(MultipleSameColumnTestRecord.class));

        assertThrows(WorkbookRecordException.class, () -> new WorkbookRecordMapper<>(MultipleSameMetadataTestRecord.class));
        assertThrows(WorkbookRecordException.class, () -> new WorkbookRecordMapper<>(MultipleSameMetadata2TestRecord.class));
        assertThrows(WorkbookRecordException.class, () -> new WorkbookRecordMapper<>(MultipleSameMetadata3TestRecord.class));
    }

    @Test
    void rangeCheck() {
        final WorkbookRecordMapper<TestRecord> mapper = new WorkbookRecordMapper<>(TestRecord.class);
        assertTrue(mapper.withinRange(1));
        assertFalse(mapper.withinRange(0));
        assertFalse(mapper.withinRange(5));
        assertTrue(mapper.beyondRange(5));
        assertFalse(mapper.beyondRange(4));

        assertTrue(mapper.withinRange(1, 2));
        assertFalse(mapper.withinRange(0, 2));
        assertFalse(mapper.withinRange(1, 10));
        assertFalse(mapper.withinRange(1, 1));

        assertTrue(mapper.withinRange(1, 2, 3));
        assertTrue(mapper.withinRange(1, 2, 1));
        assertFalse(mapper.withinRange(1, 2, 0));
        assertTrue(mapper.withinRange(1, 2, 5));
        assertFalse(mapper.withinRange(1, 2, 8));
        assertFalse(mapper.withinRange(1, 1, 1));
    }

    @Test
    void setValue() {
        final WorkbookRecordMapper<TestRecord> mapper = new WorkbookRecordMapper<>(TestRecord.class);

        final TestRecord t = new TestRecord();
        mapper.setValue(t, 0, CellValue.newInstance("a"));
        assertEquals("a", t.a);
        mapper.setValue(t, 1, CellValue.newInstance("b"));
        assertEquals("b", t.b);
        mapper.setValue(t, 2, CellValue.newInstance("c"));
        assertEquals("c", t.c);
        assertDoesNotThrow(() -> mapper.setValue(t, 3, CellValue.newInstance("d")));
    }

    @Test
    void setMetadata() {
        final WorkbookRecordMapper<TestRecord> mapper = new WorkbookRecordMapper<>(TestRecord.class);
        final TestRecord t = new TestRecord();
        mapper.setMetadata(t, WorkbookRecord.MetadataType.SHEET_INDEX, 1);
        assertEquals(1, t.sheetIndex);
        mapper.setMetadata(t, WorkbookRecord.MetadataType.SHEET_NAME, "Sheet1");
        assertEquals("Sheet1", t.sheetName);
        mapper.setMetadata(t, WorkbookRecord.MetadataType.ROW_NUMBER, 2);
        assertEquals(2, t.rowNumber);

        final WorkbookRecordMapper<TestRecord2> mapper2 = new WorkbookRecordMapper<>(TestRecord2.class);
        final TestRecord2 t2 = new TestRecord2();
        assertDoesNotThrow(() -> mapper2.setMetadata(t2, WorkbookRecord.MetadataType.SHEET_INDEX, 1));
    }

    public static class ErrorTestRecord {
        @Property(column = 0)
        private String a;
        @Property(column = 1)
        private String b;
        @Property(column = 2)
        private String c;
    }

    @WorkbookRecord(startSheet = 1, endSheet = 5,
            startRow = 2, endRow = 10, startColumn = 1, endColumn = 6)
    public static class TestRecord {
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.SHEET_INDEX)
        private Integer sheetIndex;
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.SHEET_NAME)
        private String sheetName;
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.ROW_NUMBER)
        private int rowNumber;

        @Property(column = 0)
        private String a;
        @Property(column = 1)
        private String b;
        @Property(column = 2)
        private String c;
    }

    @WorkbookRecord
    public static class TestRecord2 {
        @Property(column = 0)
        private String a;
        @Property(column = 1)
        private String b;
    }

    @WorkbookRecord
    public static class MultipleSameColumnTestRecord {
        @Property(column = 0)
        private String a;
        @Property(column = 0)
        private String b;
    }

    @WorkbookRecord
    public static class MultipleSameMetadataTestRecord {
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.SHEET_INDEX)
        private Integer sheetIndex;
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.SHEET_INDEX)
        private Integer sheetIndex2;
    }

    @WorkbookRecord
    public static class MultipleSameMetadata2TestRecord {
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.SHEET_NAME)
        private String sheetName;
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.SHEET_NAME)
        private String sheetName2;
    }

    @WorkbookRecord
    public static class MultipleSameMetadata3TestRecord {
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.ROW_NUMBER)
        private Integer rowNumber;
        @WorkbookRecord.Metadata(WorkbookRecord.MetadataType.ROW_NUMBER)
        private Integer rowNumber2;
    }

}