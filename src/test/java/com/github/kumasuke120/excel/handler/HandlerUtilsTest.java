package com.github.kumasuke120.excel.handler;


import com.github.kumasuke120.excel.CellValue;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class HandlerUtilsTest {

    @Test
    void newInstance() {
        try {
            final Constructor<HandlerUtils> constructor = HandlerUtils.class.getDeclaredConstructor();
            constructor.setAccessible(true);
            constructor.newInstance();
        } catch (InvocationTargetException e) {
            assertTrue(e.getTargetException() instanceof UnsupportedOperationException);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void ensureWorkbookRecordClass() {
        assertThrows(NullPointerException.class, () -> HandlerUtils.ensureWorkbookRecordClass(null));
        assertThrows(WorkbookRecordException.class, () -> HandlerUtils.ensureWorkbookRecordClass(String.class));
        assertDoesNotThrow(() -> HandlerUtils.ensureWorkbookRecordClass(TestRecord.class));
    }

    @Test
    void getDefaultValue() {
        assertEquals(0, HandlerUtils.getDefaultValue(int.class));
        assertEquals((byte) 0, HandlerUtils.getDefaultValue(byte.class));
        assertEquals((short) 0, HandlerUtils.getDefaultValue(short.class));
        assertEquals(0L, HandlerUtils.getDefaultValue(long.class));
        assertEquals(0.0f, HandlerUtils.getDefaultValue(float.class));
        assertEquals(0.0d, HandlerUtils.getDefaultValue(double.class));
        assertEquals(false, HandlerUtils.getDefaultValue(boolean.class));
        assertEquals('\0', HandlerUtils.getDefaultValue(char.class));
        assertNull(HandlerUtils.getDefaultValue(String.class));
        assertThrows(AssertionError.class, () -> HandlerUtils.getDefaultValue(void.class));
    }

    @Test
    void asByte() {
        assertEquals((byte) 1, HandlerUtils.asByte((byte) 1));
        assertEquals((byte) 1, HandlerUtils.asByte(1));
        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asByte("string"));
    }

    @Test
    void asShort() {
        assertEquals((short) 1, HandlerUtils.asShort((short) 1));
        assertEquals((short) 1, HandlerUtils.asShort(1));
        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asShort("string"));
    }

    @Test
    void asFloat() {
        assertEquals(1.0f, HandlerUtils.asFloat(1.0f));
        assertEquals(1.0f, HandlerUtils.asFloat(1));
        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asFloat("string"));
    }

    @Test
    void asBigInteger() {
        assertEquals(BigInteger.ONE, HandlerUtils.asBigInteger(BigInteger.ONE));
        assertEquals(BigInteger.ONE, HandlerUtils.asBigInteger(BigDecimal.ONE));
        assertEquals(BigInteger.ONE, HandlerUtils.asBigInteger(1));
        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asBigInteger("string"));
    }

    @Test
    void asDate() {
        Date now = new Date();
        assertEquals(now, HandlerUtils.asDate(now));

        final LocalDate localDateNow = LocalDate.now();
        assertEquals(Date.from(localDateNow.atStartOfDay(ZoneId.systemDefault()).toInstant()), HandlerUtils.asDate(localDateNow));

        final LocalDateTime localDateTimeNow = LocalDateTime.now();
        assertEquals(Date.from(localDateTimeNow.atZone(ZoneId.systemDefault()).toInstant()), HandlerUtils.asDate(localDateTimeNow));
        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asDate("string"));
    }

    @Test
    void asSqlTime() {
        final java.sql.Time now = new java.sql.Time(new Date().getTime());
        assertEquals(now, HandlerUtils.asSqlTime(now));

        final LocalTime localTimeNow = LocalTime.now();
        assertEquals(java.sql.Time.valueOf(localTimeNow), HandlerUtils.asSqlTime(localTimeNow));
        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asSqlTime(localTimeNow.toString()));
        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asSqlTime("string"));

        final LocalDateTime localDateTimeNow = LocalDateTime.now();
        assertEquals(java.sql.Time.valueOf(localDateTimeNow.toLocalTime()), HandlerUtils.asSqlTime(localDateTimeNow));
    }

    @Test
    void asSqlDate() {
        final java.sql.Date now = new java.sql.Date(new Date().getTime());
        assertEquals(now, HandlerUtils.asSqlDate(now));
        assertEquals(java.sql.Date.valueOf(LocalDate.now()), HandlerUtils.asSqlDate(LocalDate.now()));
        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asSqlDate("string"));

        final LocalDateTime localDateTimeNow = LocalDateTime.now();
        assertEquals(java.sql.Date.valueOf(localDateTimeNow.toLocalDate()), HandlerUtils.asSqlDate(localDateTimeNow));

        final Date date = new Date();
        assertEquals(new java.sql.Date(date.getTime()), HandlerUtils.asSqlDate(date));
    }

    @Test
    void asSqlTimestamp() {
        final java.sql.Timestamp now = new java.sql.Timestamp(new Date().getTime());
        assertEquals(now, HandlerUtils.asSqlTimestamp(now));
        final LocalDateTime localDateTimeNow = LocalDateTime.now();
        assertEquals(java.sql.Timestamp.valueOf(localDateTimeNow), HandlerUtils.asSqlTimestamp(localDateTimeNow));

        final LocalDate localDateNow = LocalDate.now();
        assertEquals(java.sql.Timestamp.valueOf(localDateNow.atStartOfDay()), HandlerUtils.asSqlTimestamp(localDateNow));

        assertThrows(IllegalArgumentException.class, () -> HandlerUtils.asSqlTimestamp("string"));

        final Date date = new Date();
        assertEquals(new java.sql.Timestamp(date.getTime()), HandlerUtils.asSqlTimestamp(date));
    }

    @Test
    void primitiveToWrapper() {
        assertEquals(Boolean.class, HandlerUtils.primitiveToWrapper(boolean.class));
        assertEquals(Byte.class, HandlerUtils.primitiveToWrapper(byte.class));
        assertEquals(Short.class, HandlerUtils.primitiveToWrapper(short.class));
        assertEquals(Integer.class, HandlerUtils.primitiveToWrapper(int.class));
        assertEquals(Long.class, HandlerUtils.primitiveToWrapper(long.class));
        assertEquals(Float.class, HandlerUtils.primitiveToWrapper(float.class));
        assertEquals(Double.class, HandlerUtils.primitiveToWrapper(double.class));
        assertEquals(Character.class, HandlerUtils.primitiveToWrapper(char.class));
        assertEquals(Void.class, HandlerUtils.primitiveToWrapper(void.class));
        assertEquals(String.class, HandlerUtils.primitiveToWrapper(String.class));
        assertNotEquals(String.class, HandlerUtils.primitiveToWrapper(List.class));
    }

    @Test
    void isValueEmpty() {
        assertTrue(HandlerUtils.isValueEmpty(null));
        assertTrue(HandlerUtils.isValueEmpty(CellValue.newInstance(null)));
        assertTrue(HandlerUtils.isValueEmpty(CellValue.newInstance("")));
        assertTrue(HandlerUtils.isValueEmpty(CellValue.newInstance("   ")));
        assertFalse(HandlerUtils.isValueEmpty(CellValue.newInstance("1")));
    }

    @WorkbookRecord
    public static class TestRecord {
    }

}