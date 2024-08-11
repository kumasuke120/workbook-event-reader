package com.github.kumasuke120.excel;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.Objects;

import static org.junit.jupiter.api.Assertions.*;

// we only test for missed branches here because other test classes will test all normal branches
class UtilTest {

    @Test
    void isAWholeNumber() {
        assertFalse(Util.isAWholeNumber(null));
    }

    @Test
    void isADecimalFraction() {
        assertFalse(Util.isADecimalFraction(null));
        assertFalse(Util.isADecimalFraction("null"));
    }

    @Test
    void decimalStringToDecimal() {
        assertNull(Util.decimalStringToDecimal(null));
        assertEquals("null", Util.decimalStringToDecimal("null"));
        assertEquals(1234L, Util.decimalStringToDecimal("1234"));
        assertEquals(1234L, Util.decimalStringToDecimal("1,234"));
        assertEquals(1234.56, Util.decimalStringToDecimal("1,234.56"));
    }

    @Test
    void toInt() {
        assertEquals(0, Util.toInt(null, 0));
        assertEquals(0, Util.toInt("null", 0));
    }

    @Test
    void isValidExcelDate() {
        assertFalse(Util.isValidExcelDate(Double.NaN));
        assertFalse(Util.isValidExcelDate(-1));
        assertFalse(Util.isValidExcelDate(2958466));
        assertTrue(Util.isValidExcelDate(2));
        assertTrue(Util.isValidExcelDate(0));
    }

    @Test
    void toJsr310DateOrTime() {
        assertNull(Util.toJsr310DateOrTime(2958466, false));
    }

    @Test
    void cellReferenceToRowAndColumn() {
        assertNull(Util.cellReferenceToRowAndColumn("ox"));
        assertEquals(26, Objects.requireNonNull(Util.cellReferenceToRowAndColumn("AA1")).getValue());
    }

    @Test
    void isATextFormat() {
        assertTrue(Util.isATextFormat(0x31, null));
        assertTrue(Util.isATextFormat(-1, "@"));
        assertFalse(Util.isATextFormat(-1, null));
        assertFalse(Util.isATextFormat(-1, "General"));
    }

    @Test
    void newInstance() {
        try {
            final Constructor<Util> constructor = Util.class.getDeclaredConstructor();
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
        assertEquals(Integer.MAX_VALUE + 1L, Util.toRelativeType(Integer.MAX_VALUE + 1d));
        assertEquals(Integer.MIN_VALUE - 1L, Util.toRelativeType(Integer.MIN_VALUE - 1d));
        assertEquals(1, Util.toRelativeType(1d));
    }

}