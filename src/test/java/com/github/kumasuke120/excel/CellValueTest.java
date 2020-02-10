package com.github.kumasuke120.excel;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.*;

class CellValueTest {

    @Test
    @DisplayName("null()")
    void _null() {
        final var cellValue = newCellValue(null);

        assertTrue(cellValue.isNull());
        assertNull(cellValue.originalValue());
        assertThrows(NullPointerException.class, cellValue::originalType);
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertThrows(CellValueCastException.class, cellValue::stringValue);
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);
    }

    @Test
    @DisplayName("boolean()")
    void _boolean() {
        final var cellValue = newCellValue(false);

        assertFalse(cellValue.isNull());
        assertEquals(false, cellValue.originalValue());
        assertEquals(Boolean.class, cellValue.originalType());
        assertDoesNotThrow(() -> {
            final boolean booleanValue = cellValue.booleanValue();
            assertFalse(booleanValue);
        });
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals(Boolean.toString(false), stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final var cellValue2 = newCellValue(0);
        assertFalse(cellValue2.isNull());
        assertDoesNotThrow(() -> {
            final boolean booleanValue = cellValue2.booleanValue();
            assertFalse(booleanValue);
        });
    }

    @Test
    @DisplayName("int()")
    void _int() {
        final var cellValue = newCellValue(1);

        assertFalse(cellValue.isNull());
        assertEquals(1, cellValue.originalValue());
        assertEquals(Integer.class, cellValue.originalType());
        assertDoesNotThrow(() -> {
            final boolean booleanValue = cellValue.booleanValue();
            assertTrue(booleanValue);
        });
        assertDoesNotThrow(() -> {
            final int intValue = cellValue.intValue();
            assertEquals(1, intValue);
        });
        assertDoesNotThrow(() -> {
            final long longValue = cellValue.longValue();
            assertEquals(1L, longValue);
        });
        assertDoesNotThrow(() -> {
            final double doubleValue = cellValue.doubleValue();
            assertEquals(1D, doubleValue);
        });
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals("1", stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);


        final var cellValue2 = newCellValue("1");
        assertEquals(String.class, cellValue2.originalType());
        assertDoesNotThrow(() -> {
            final int intValue = cellValue2.intValue();
            assertEquals(1, intValue);
        });

        final var cellValue3 = newCellValue("1.20");
        assertEquals(String.class, cellValue3.originalType());
        assertDoesNotThrow(() -> {
            final int intValue = cellValue3.intValue();
            assertEquals(1, intValue);
        });
    }

    @Test
    @DisplayName("long()")
    void _long() {
        final var cellValue = newCellValue(1L);

        assertFalse(cellValue.isNull());
        assertEquals(1L, cellValue.originalValue());
        assertEquals(Long.class, cellValue.originalType());
        assertDoesNotThrow(() -> {
            final boolean booleanValue = cellValue.booleanValue();
            assertTrue(booleanValue);
        });
        assertDoesNotThrow(() -> {
            final int intValue = cellValue.intValue();
            assertEquals(1, intValue);
        });
        assertDoesNotThrow(() -> {
            final long longValue = cellValue.longValue();
            assertEquals(1L, longValue);
        });
        assertDoesNotThrow(() -> {
            final double doubleValue = cellValue.doubleValue();
            assertEquals(1D, doubleValue);
        });
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals("1", stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final var cellValue2 = newCellValue(Long.toString(Long.MAX_VALUE));
        assertEquals(String.class, cellValue2.originalType());
        assertDoesNotThrow(() -> {
            final long longValue = cellValue2.longValue();
            assertEquals(Long.MAX_VALUE, longValue);
        });

        final var cellValue3 = newCellValue("1.20");
        assertEquals(String.class, cellValue3.originalType());
        assertDoesNotThrow(() -> {
            final long longValue = cellValue3.longValue();
            assertEquals(1L, longValue);
        });
    }

    @Test
    @DisplayName("double()")
    void _double() {
        final var cellValue = newCellValue(1D);

        assertFalse(cellValue.isNull());
        assertEquals(1D, cellValue.originalValue());
        assertEquals(Double.class, cellValue.originalType());
        assertDoesNotThrow(() -> {
            final boolean booleanValue = cellValue.booleanValue();
            assertTrue(booleanValue);
        });
        assertDoesNotThrow(() -> {
            final int intValue = cellValue.intValue();
            assertEquals(1, intValue);
        });
        assertDoesNotThrow(() -> {
            final long longValue = cellValue.longValue();
            assertEquals(1L, longValue);
        });
        assertDoesNotThrow(() -> {
            final double doubleValue = cellValue.doubleValue();
            assertEquals(1D, doubleValue);
        });
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals(Double.toString(1D), stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final var cellValue2 = newCellValue("1.20");
        assertEquals(String.class, cellValue2.originalType());
        assertDoesNotThrow(() -> {
            final double doubleValue = cellValue2.doubleValue();
            assertEquals(1.2, doubleValue);
        });
    }

    @Test
    void string() {
        final var cellValue = newCellValue("1");

        assertFalse(cellValue.isNull());
        assertEquals("1", cellValue.originalValue());
        assertEquals(String.class, cellValue.originalType());
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertDoesNotThrow(() -> {
            final int intValue = cellValue.intValue();
            assertEquals(1, intValue);
        });
        assertDoesNotThrow(() -> {
            final long longValue = cellValue.longValue();
            assertEquals(1L, longValue);
        });
        assertDoesNotThrow(() -> {
            final double doubleValue = cellValue.doubleValue();
            assertEquals(1D, doubleValue);
        });
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals("1", stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);
    }

    @Test
    void localTime() {
        final var cellValue = newCellValue("12:34:56");

        assertFalse(cellValue.isNull());
        assertEquals("12:34:56", cellValue.originalValue());
        assertEquals(String.class, cellValue.originalType());
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertDoesNotThrow(() -> {
            final LocalTime localTimeValue = cellValue.localTimeValue();
            assertEquals(LocalTime.of(12, 34, 56), localTimeValue);

            final DateTimeFormatter theFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");
            final LocalTime localTimeValue2 = cellValue.localTimeValue(theFormatter);
            assertEquals(LocalTime.of(12, 34, 56), localTimeValue2);

            final LocalTime localTimeValue3 = cellValue.localTimeValue(Set.of(theFormatter));
            assertEquals(LocalTime.of(12, 34, 56), localTimeValue3);
        });
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final LocalTime theTimeValue = LocalTime.now();
        final var cellValue2 = newCellValue(theTimeValue);
        assertEquals(theTimeValue, cellValue2.originalValue());
        assertEquals(LocalTime.class, cellValue2.originalType());
        assertEquals(theTimeValue, cellValue2.localTimeValue());
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);
    }

    @Test
    void localDate() {
        final var cellValue = newCellValue("2018-11-11");

        assertFalse(cellValue.isNull());
        assertEquals("2018-11-11", cellValue.originalValue());
        assertEquals(String.class, cellValue.originalType());
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals("2018-11-11", stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertDoesNotThrow(() -> {
            final LocalDate localDateValue = cellValue.localDateValue();
            assertEquals(LocalDate.of(2018, 11, 11), localDateValue);

            final DateTimeFormatter theFormatter = DateTimeFormatter.ofPattern("uuuu-MM-dd");
            final LocalDate localDateValue2 = cellValue.localDateValue(theFormatter);
            assertEquals(LocalDate.of(2018, 11, 11), localDateValue2);

            final LocalDate localDateValue3 = cellValue.localDateValue(Set.of(theFormatter));
            assertEquals(LocalDate.of(2018, 11, 11), localDateValue3);
        });
        assertDoesNotThrow(() -> {
            final LocalDateTime localDateTimeValue = cellValue.localDateTimeValue();
            assertEquals(LocalDateTime.of(2018, 11, 11, 0, 0, 0),
                         localDateTimeValue);

            final DateTimeFormatter theFormatter = DateTimeFormatter.ofPattern("uuuu-MM-dd");
            final LocalDateTime localDateTimeValue2 = cellValue.localDateTimeValue(theFormatter);
            assertEquals(LocalDate.of(2018, 11, 11).atStartOfDay(), localDateTimeValue2);

            final LocalDateTime localDateTimeValue3 = cellValue.localDateTimeValue(Set.of(theFormatter));
            assertEquals(LocalDate.of(2018, 11, 11).atStartOfDay(), localDateTimeValue3);
        });

        final LocalDate theDateValue = LocalDate.now();
        final var cellValue2 = newCellValue(theDateValue);
        assertEquals(theDateValue, cellValue2.originalValue());
        assertEquals(LocalDate.class, cellValue2.originalType());
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertEquals(theDateValue, cellValue2.localDateValue());
        assertEquals(theDateValue.atStartOfDay(), cellValue2.localDateTimeValue());
    }

    @Test
    void localDateTime() {
        final var cellValue = newCellValue("2011-11-11T11:11:11");

        assertFalse(cellValue.isNull());
        assertEquals("2011-11-11T11:11:11", cellValue.originalValue());
        assertEquals(String.class, cellValue.originalType());
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals("2011-11-11T11:11:11", stringValue);
        });
        assertDoesNotThrow(() -> {
            final LocalTime localTimeValue = cellValue.localTimeValue();
            assertEquals(LocalTime.of(11, 11, 11), localTimeValue);
        });
        assertDoesNotThrow(() -> {
            final LocalDate localDateValue = cellValue.localDateValue();
            assertEquals(LocalDate.of(2011, 11, 11), localDateValue);

            final DateTimeFormatter theFormatter = DateTimeFormatter.ofPattern("uuuu-MM-dd'T'HH:mm:ss");
            final LocalDate localDateValue2 = cellValue.localDateValue(theFormatter);
            assertEquals(LocalDate.of(2011, 11, 11), localDateValue2);

            final LocalDate localDateValue3 = cellValue.localDateValue(Set.of(theFormatter));
            assertEquals(LocalDate.of(2011, 11, 11), localDateValue3);
        });
        assertDoesNotThrow(() -> {
            final LocalDateTime localDateTimeValue = cellValue.localDateTimeValue();
            assertEquals(LocalDateTime.of(2011, 11, 11, 11, 11, 11),
                         localDateTimeValue);
            final DateTimeFormatter theFormatter = DateTimeFormatter.ofPattern("uuuu-MM-dd'T'HH:mm:ss");
            final LocalDateTime localDateTimeValue2 = cellValue.localDateTimeValue(theFormatter);
            assertEquals(LocalDateTime.of(2011, 11, 11, 11, 11, 11),
                         localDateTimeValue2);

            final LocalDateTime localDateTimeValue3 = cellValue.localDateTimeValue(Set.of(theFormatter));
            assertEquals(LocalDateTime.of(2011, 11, 11, 11, 11, 11),
                         localDateTimeValue3);
        });

        final LocalDateTime theDateTimeValue = LocalDateTime.now();
        final var cellValue2 = newCellValue(theDateTimeValue);
        assertEquals(theDateTimeValue, cellValue2.originalValue());
        assertEquals(LocalDateTime.class, cellValue2.originalType());
        assertEquals(theDateTimeValue.toLocalTime(), cellValue2.localTimeValue());
        assertEquals(theDateTimeValue.toLocalDate(), cellValue2.localDateValue());
        assertEquals(theDateTimeValue, cellValue2.localDateTimeValue());
    }

    @Test
    void equalsAndHashCode() {
        // ide generated code, test for coverage
        final var a = newCellValue(1);
        final var b = newCellValue(1);

        assertEquals(a, b);
        assertEquals(a.hashCode(), b.hashCode());
    }

    @Test
    @DisplayName("toString()")
    void _toString() {
        final var a = newCellValue(1);
        final var b = newCellValue(null);

        assertTrue(a.toString().matches("app\\.kumasuke\\.excel\\.CellValue" +
                                                "\\{type = java\\.lang\\.Integer, value = 1}" +
                                                "@[a-z0-9]+"));
        assertTrue(b.toString().matches("app\\.kumasuke\\.excel\\.CellValue" +
                                                "\\{value = null}" +
                                                "@[a-z0-9]+"));
    }

    @Test
    void newInstance() {
        final CellValue a = newCellValueByStaticMethod(null);
        final CellValue b = newCellValueByStaticMethod(null);
        final CellValue c = newCellValue(null);

        assertNotNull(a);
        assertNotNull(b);
        assertNotNull(c);

        assertSame(a, b);
        assertSame(a, newCellValueByStaticMethod(null));
        assertNotSame(a, c);
        assertNotSame(c, newCellValueByStaticMethod(null));
    }

    private CellValue newCellValue(Object originalValue) {
        try {
            final Constructor<CellValue> constructor = CellValue.class.getDeclaredConstructor(Object.class);
            constructor.setAccessible(true);
            return constructor.newInstance(originalValue);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    @SuppressWarnings("SameParameterValue")
    private CellValue newCellValueByStaticMethod(Object originalValue) {
        try {
            final Method method = CellValue.class.getDeclaredMethod("newInstance", Object.class);
            method.setAccessible(true);
            return (CellValue) method.invoke(null, originalValue);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

}
