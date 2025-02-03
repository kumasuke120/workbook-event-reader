package com.github.kumasuke120.excel;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.Objects;

import static org.junit.jupiter.api.Assertions.*;

// we only test for missed branches here because other test classes will test all normal branches
class ReaderUtilsTest {

    @Test
    void isAWholeNumber() {
        assertFalse(ReaderUtils.isAWholeNumber(null));
    }

    @Test
    void isADecimalFraction() {
        assertFalse(ReaderUtils.isADecimalFraction(null));
        assertFalse(ReaderUtils.isADecimalFraction("null"));
    }

    @Test
    void decimalStringToDecimal() {
        assertNull(ReaderUtils.decimalStringToDecimal(null));
        assertEquals("null", ReaderUtils.decimalStringToDecimal("null"));
        assertEquals(1234L, ReaderUtils.decimalStringToDecimal("1234"));
        assertEquals(1234L, ReaderUtils.decimalStringToDecimal("1,234"));
        assertEquals(1234.56, ReaderUtils.decimalStringToDecimal("1,234.56"));
        assertEquals(new BigDecimal("12345678.9"), ReaderUtils.decimalStringToDecimal("12,345,678.9"));
    }

    @Test
    void toInt() {
        assertEquals(0, ReaderUtils.toInt(null, 0));
        assertEquals(0, ReaderUtils.toInt("null", 0));
    }

    @Test
    void isValidExcelDate() {
        assertFalse(ReaderUtils.isValidExcelDate(Double.NaN));
        assertFalse(ReaderUtils.isValidExcelDate(-1));
        assertFalse(ReaderUtils.isValidExcelDate(2958466));
        assertTrue(ReaderUtils.isValidExcelDate(2));
        assertTrue(ReaderUtils.isValidExcelDate(0));
    }

    @Test
    void toJsr310DateOrTime() {
        assertNull(ReaderUtils.toJsr310DateOrTime(2958466, false));
    }

    @Test
    void cellReferenceToRowAndColumn() {
        assertNull(ReaderUtils.cellReferenceToRowAndColumn("ox"));
        assertEquals(26, Objects.requireNonNull(ReaderUtils.cellReferenceToRowAndColumn("AA1")).getValue());
    }

    @Test
    void isATextFormat() {
        assertTrue(ReaderUtils.isATextFormat(0x31, null));
        assertTrue(ReaderUtils.isATextFormat(-1, "@"));
        assertFalse(ReaderUtils.isATextFormat(-1, null));
        assertFalse(ReaderUtils.isATextFormat(-1, "General"));
    }

    @Test
    void newInstance() {
        try {
            final Constructor<ReaderUtils> constructor = ReaderUtils.class.getDeclaredConstructor();
            constructor.setAccessible(true);
            constructor.newInstance();
        } catch (InvocationTargetException e) {
            assertTrue(e.getTargetException() instanceof UnsupportedOperationException);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void toRelativeType() {
        assertEquals(Integer.MAX_VALUE + 1L, ReaderUtils.toRelativeType(Integer.MAX_VALUE + 1d));
        assertEquals(Integer.MIN_VALUE - 1L, ReaderUtils.toRelativeType(Integer.MIN_VALUE - 1d));
        assertEquals(1, ReaderUtils.toRelativeType(1d));
    }

}