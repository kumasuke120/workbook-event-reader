package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.handler.WorkbookRecord.CellValueType;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class CellValueTypeTest {

    @Test
    void getAutoValue() {
        final CellValue cellValue = newCellValue(null);
        assertThrows(UnsupportedOperationException.class, () -> CellValueType.AUTO.getValue(cellValue));
    }

    @Test
    void getBooleanValue() {
        final CellValue cellValue = newCellValue(true);
        assertEquals(true, CellValueType.BOOLEAN.getValue(cellValue));
    }

    @Test
    void getIntegerValue() {
        final CellValue cellValue = newCellValue(123);
        assertEquals(123, CellValueType.INTEGER.getValue(cellValue));
    }

    @Test
    void getLongValue() {
        final CellValue cellValue = newCellValue(123L);
        assertEquals(123L, CellValueType.LONG.getValue(cellValue));
    }

    @Test
    void getDoubleValue() {
        final CellValue cellValue = newCellValue(123.45);
        assertEquals(123.45, CellValueType.DOUBLE.getValue(cellValue));
    }

    @Test
    void getDecimalValue() {
        final CellValue cellValue = newCellValue(new BigDecimal("123.45"));
        assertEquals(new BigDecimal("123.45"), CellValueType.DECIMAL.getValue(cellValue));
    }

    @Test
    void getStringValue() {
        final CellValue cellValue = newCellValue("test");
        assertEquals("test", CellValueType.STRING.getValue(cellValue));
    }

    @Test
    void getTimeValue() {
        final LocalTime time = LocalTime.of(12, 34, 56);
        final CellValue cellValue = newCellValue(time);
        assertEquals(time, CellValueType.TIME.getValue(cellValue));
    }

    @Test
    void getDateValue() {
        final LocalDate date = LocalDate.of(2023, 10, 1);
        final CellValue cellValue = newCellValue(date);
        assertEquals(date, CellValueType.DATE.getValue(cellValue));
    }

    @Test
    void getDateTimeValue() {
        final LocalDateTime dateTime = LocalDateTime.of(2023, 10, 1, 12, 34, 56);
        final CellValue cellValue = newCellValue(dateTime);
        assertEquals(dateTime, CellValueType.DATETIME.getValue(cellValue));
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

}
