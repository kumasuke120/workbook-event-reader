package com.github.kumasuke120.excel;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Collections;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.*;

class LenientCellValueTest {

    @Test
    @DisplayName("null()")
    void _null() {
        final LenientCellValue cellValue = newLenientCellValue(null);

        assertTrue(cellValue.isNull());
        assertNull(cellValue.originalValue());
        assertThrows(NullPointerException.class, cellValue::originalType);
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertThrows(CellValueCastException.class, cellValue::bigDecimalValue);
        assertThrows(CellValueCastException.class, cellValue::stringValue);
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);
    }

    @Test
    @DisplayName("boolean()")
    void _boolean() {
        final LenientCellValue cellValue = newLenientCellValue(false);

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
        assertThrows(CellValueCastException.class, cellValue::bigDecimalValue);
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals(Boolean.toString(false), stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final LenientCellValue cellValue2 = newLenientCellValue(0);
        assertFalse(cellValue2.isNull());
        assertDoesNotThrow(() -> {
            final boolean booleanValue = cellValue2.booleanValue();
            assertFalse(booleanValue);
        });
    }

    @Test
    @DisplayName("int()")
    void _int() {
        final LenientCellValue cellValue = newLenientCellValue(1);

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
            final BigDecimal doubleValue = cellValue.bigDecimalValue();
            assertEquals(BigDecimal.ONE, doubleValue);
        });
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals("1", stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);


        final LenientCellValue cellValue2 = newLenientCellValue("1");
        assertEquals(String.class, cellValue2.originalType());
        assertDoesNotThrow(() -> {
            final int intValue = cellValue2.intValue();
            assertEquals(1, intValue);
        });

        final LenientCellValue cellValue3 = newLenientCellValue("1.20");
        assertEquals(String.class, cellValue3.originalType());
        assertDoesNotThrow(() -> {
            final int intValue = cellValue3.intValue();
            assertEquals(1, intValue);
        });
    }

    @Test
    @DisplayName("long()")
    void _long() {
        final LenientCellValue cellValue = newLenientCellValue(1L);

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
            final BigDecimal doubleValue = cellValue.bigDecimalValue();
            assertEquals(BigDecimal.ONE, doubleValue);
        });
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals("1", stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final LenientCellValue cellValue2 = newLenientCellValue(Long.toString(Long.MAX_VALUE));
        assertEquals(String.class, cellValue2.originalType());
        assertDoesNotThrow(() -> {
            final long longValue = cellValue2.longValue();
            assertEquals(Long.MAX_VALUE, longValue);
        });

        final LenientCellValue cellValue3 = newLenientCellValue("1.20");
        assertEquals(String.class, cellValue3.originalType());
        assertDoesNotThrow(() -> {
            final long longValue = cellValue3.longValue();
            assertEquals(1L, longValue);
        });
    }

    @SuppressWarnings("SimplifiableAssertion")
    @Test
    @DisplayName("double()")
    void _double() {
        final LenientCellValue cellValue = newLenientCellValue(1D);

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
            final BigDecimal bigDecimalValue = cellValue.bigDecimalValue();
            assertTrue(BigDecimal.ONE.compareTo(bigDecimalValue) == 0);
        });
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals(Double.toString(1D), stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final LenientCellValue cellValue2 = newLenientCellValue("1.20");
        assertEquals(String.class, cellValue2.originalType());
        assertDoesNotThrow(() -> {
            final double doubleValue = cellValue2.doubleValue();
            assertEquals(1.2, doubleValue);
        });
    }

    @SuppressWarnings("SimplifiableAssertion")
    @Test
    void bigDecimal() {
        final LenientCellValue cellValue = newLenientCellValue(BigDecimal.ONE);

        assertFalse(cellValue.isNull());
        assertEquals(BigDecimal.ONE, cellValue.originalValue());
        assertEquals(BigDecimal.class, cellValue.originalType());
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
            final BigDecimal bigDecimalValue = cellValue.bigDecimalValue();
            assertTrue(BigDecimal.ONE.compareTo(bigDecimalValue) == 0);
        });
        assertDoesNotThrow(() -> {
            final String stringValue = cellValue.stringValue();
            assertEquals("1", stringValue);
        });
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final LenientCellValue cellValue2 = newLenientCellValue("1.20");
        assertEquals(String.class, cellValue2.originalType());
        assertDoesNotThrow(() -> {
            final BigDecimal decimalValue = cellValue2.bigDecimalValue();
            assertEquals(new BigDecimal("1.20"), decimalValue);
        });
    }

    @Test
    void string() {
        final LenientCellValue cellValue = newLenientCellValue("1");

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
            final BigDecimal doubleValue = cellValue.bigDecimalValue();
            assertEquals(BigDecimal.ONE, doubleValue);
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
        final LenientCellValue cellValue = newLenientCellValue("12:34:56");

        assertFalse(cellValue.isNull());
        assertEquals("12:34:56", cellValue.originalValue());
        assertEquals(String.class, cellValue.originalType());
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertThrows(CellValueCastException.class, cellValue::bigDecimalValue);
        assertDoesNotThrow(() -> {
            final LocalTime localTimeValue = cellValue.localTimeValue();
            assertEquals(LocalTime.of(12, 34, 56), localTimeValue);

            final DateTimeFormatter theFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");
            final LocalTime localTimeValue2 = cellValue.localTimeValue(theFormatter);
            assertEquals(LocalTime.of(12, 34, 56), localTimeValue2);

            final LocalTime localTimeValue3 = cellValue.localTimeValue(Collections.singleton(theFormatter));
            assertEquals(LocalTime.of(12, 34, 56), localTimeValue3);
        });
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final LocalTime theTimeValue = LocalTime.now();
        final LenientCellValue cellValue2 = newLenientCellValue(theTimeValue);
        assertEquals(theTimeValue, cellValue2.originalValue());
        assertEquals(LocalTime.class, cellValue2.originalType());
        assertEquals(theTimeValue, cellValue2.localTimeValue());
        assertThrows(CellValueCastException.class, cellValue::localDateValue);
        assertThrows(CellValueCastException.class, cellValue::localDateTimeValue);

        final LenientCellValue cellValue3 = newLenientCellValue("2020-01-20T12:34:56");
        assertFalse(cellValue3.isNull());
        assertEquals("2020-01-20T12:34:56", cellValue3.originalValue());
        assertEquals(String.class, cellValue3.originalType());
        assertEquals(LocalTime.of(12, 34, 56), cellValue3.localTimeValue());
    }

    @Test
    void localDate() {
        final LenientCellValue cellValue = newLenientCellValue("2018-11-11");

        assertFalse(cellValue.isNull());
        assertEquals("2018-11-11", cellValue.originalValue());
        assertEquals(String.class, cellValue.originalType());
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertThrows(CellValueCastException.class, cellValue::bigDecimalValue);
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

            final LocalDate localDateValue3 = cellValue.localDateValue(Collections.singleton(theFormatter));
            assertEquals(LocalDate.of(2018, 11, 11), localDateValue3);
        });
        assertDoesNotThrow(() -> {
            final LocalDateTime localDateTimeValue = cellValue.localDateTimeValue();
            assertEquals(LocalDateTime.of(2018, 11, 11, 0, 0, 0),
                         localDateTimeValue);

            final DateTimeFormatter theFormatter = DateTimeFormatter.ofPattern("uuuu-MM-dd");
            final LocalDateTime localDateTimeValue2 = cellValue.localDateTimeValue(theFormatter);
            assertEquals(LocalDate.of(2018, 11, 11).atStartOfDay(), localDateTimeValue2);

            final LocalDateTime localDateTimeValue3 = cellValue.localDateTimeValue(Collections.singleton(theFormatter));
            assertEquals(LocalDate.of(2018, 11, 11).atStartOfDay(), localDateTimeValue3);
        });

        final LocalDate theDateValue = LocalDate.now();
        final LenientCellValue cellValue2 = newLenientCellValue(theDateValue);
        assertEquals(theDateValue, cellValue2.originalValue());
        assertEquals(LocalDate.class, cellValue2.originalType());
        assertThrows(CellValueCastException.class, cellValue::localTimeValue);
        assertEquals(theDateValue, cellValue2.localDateValue());
        assertEquals(theDateValue.atStartOfDay(), cellValue2.localDateTimeValue());
    }

    @Test
    void localDateTime() {
        final LenientCellValue cellValue = newLenientCellValue("2011-11-11T11:11:11");

        assertFalse(cellValue.isNull());
        assertEquals("2011-11-11T11:11:11", cellValue.originalValue());
        assertEquals(String.class, cellValue.originalType());
        assertThrows(CellValueCastException.class, cellValue::booleanValue);
        assertThrows(CellValueCastException.class, cellValue::intValue);
        assertThrows(CellValueCastException.class, cellValue::longValue);
        assertThrows(CellValueCastException.class, cellValue::doubleValue);
        assertThrows(CellValueCastException.class, cellValue::bigDecimalValue);
        final CellValueCastException valueCastException =
                assertThrows(CellValueCastException.class,
                             () -> cellValue.localDateTimeValue(Collections.singletonList(null)));
        assertTrue(valueCastException.getCause() instanceof NullPointerException);
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

            final LocalDate localDateValue3 = cellValue.localDateValue(Collections.singleton(theFormatter));
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

            final LocalDateTime localDateTimeValue3 = cellValue.localDateTimeValue(Collections.singleton(theFormatter));
            assertEquals(LocalDateTime.of(2011, 11, 11, 11, 11, 11),
                         localDateTimeValue3);
        });

        final LocalDateTime theDateTimeValue = LocalDateTime.now();
        final LenientCellValue cellValue2 = newLenientCellValue(theDateTimeValue);
        assertEquals(theDateTimeValue, cellValue2.originalValue());
        assertEquals(LocalDateTime.class, cellValue2.originalType());
        assertEquals(theDateTimeValue.toLocalTime(), cellValue2.localTimeValue());
        assertEquals(theDateTimeValue.toLocalDate(), cellValue2.localDateValue());
        assertEquals(theDateTimeValue, cellValue2.localDateTimeValue());
    }

    @Test
    void mapOriginalValue() {
        // region square
        final Function<Object, Object> mappingFunction1 = (Object v) -> {
            if (v instanceof Integer) {
                final int integerValue = (int) v;
                return integerValue * integerValue;
            } else {
                return v;
            }
        };

        final LenientCellValue cellValue1 = LenientCellValue.newInstance(42);
        final LenientCellValue cellValue2 = (LenientCellValue) cellValue1.mapOriginalValue(mappingFunction1);
        assertEquals(42 * 42, cellValue2.originalValue());

        final LenientCellValue cellValue3 = LenientCellValue.newInstance("42");
        final LenientCellValue cellValue4 = (LenientCellValue) cellValue3.mapOriginalValue(mappingFunction1);
        assertSame(cellValue3, cellValue4);
        // endregion

        // region throws in function
        final Function<Object, Object> mappingFunction2 = (Object v) -> {
            throw new UnsupportedOperationException();
        };
        assertThrows(CellValueCastException.class, () -> cellValue1.mapOriginalValue(mappingFunction2));
        // endregion

        // region throws after return
        final Function<Object, Object> mappingFunction3 = (Object v) -> new UnsupportedOperationException();
        assertThrows(CellValueCastException.class, () -> cellValue1.mapOriginalValue(mappingFunction3));
        // endregion
    }

    @Test
    void trim() {
        final LenientCellValue cellValue1 = LenientCellValue.newInstance(" 1 ");
        assertEquals("1", cellValue1.trim().originalValue());

        final LenientCellValue cellValue2 = LenientCellValue.newInstance(null);
        assertSame(cellValue2, cellValue2.trim());

        final LenientCellValue cellValue3 = LenientCellValue.newInstance("1");
        assertEquals("1", cellValue3.trim().originalValue());
    }

    @Test
    void equalsAndHashCode() {
        // ide generated code, test for coverage
        final LenientCellValue a = newLenientCellValue(1);
        final LenientCellValue b = newLenientCellValue(1);

        assertEquals(a, b);
        assertEquals(a.hashCode(), b.hashCode());
    }

    @Test
    @DisplayName("toString()")
    void _toString() {
        final LenientCellValue a = newLenientCellValue(1);
        final LenientCellValue b = newLenientCellValue(null);

        assertTrue(a.toString().matches("com\\.github\\.kumasuke120\\.excel\\.LenientCellValue" +
                                                "\\{type = java\\.lang\\.Integer, value = 1}" +
                                                "@[a-z0-9]+"));
        assertTrue(b.toString().matches("com\\.github\\.kumasuke120\\.excel\\.LenientCellValue" +
                                                "\\{value = null}" +
                                                "@[a-z0-9]+"));
    }

    @Test
    void newInstance() {
        final LenientCellValue a = LenientCellValue.newInstance(null);
        final LenientCellValue b = LenientCellValue.newInstance(null);
        final LenientCellValue c = newLenientCellValue(null);

        assertNotNull(a);
        assertNotNull(b);
        assertNotNull(c);

        assertSame(a, b);
        assertSame(a, LenientCellValue.newInstance(null));
        assertNotSame(a, c);
        assertNotSame(c, LenientCellValue.newInstance(null));
    }

    @NotNull
    private LenientCellValue newLenientCellValue(@Nullable Object originalValue) {
        try {
            final Constructor<LenientCellValue> constructor = LenientCellValue.class.getDeclaredConstructor(Object.class);
            constructor.setAccessible(true);
            return constructor.newInstance(originalValue);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void strict() {
        CellValue cellValue = LenientCellValue.newInstance(1);
        assertTrue(cellValue.strict() instanceof StrictCellValue);
        assertNotSame(cellValue, cellValue.strict());
        assertNotEquals(cellValue, cellValue.strict());

        CellValue strictCellValue = cellValue.strict();
        assertDoesNotThrow(() -> {
            final int intValue = strictCellValue.intValue();
            assertEquals(1, intValue);
        });
        assertThrows(CellValueCastException.class, strictCellValue::longValue);
    }

    @Test
    void lenient() {
        CellValue cellValue = LenientCellValue.newInstance(1);
        assertTrue(cellValue.lenient() instanceof LenientCellValue);
        assertSame(cellValue, cellValue.lenient());
        assertEquals(cellValue, cellValue.lenient());

        CellValue lenientCellValue = cellValue.lenient();
        assertDoesNotThrow(() -> {
            final int intValue = lenientCellValue.intValue();
            assertEquals(1, intValue);
        });
        assertDoesNotThrow(() -> {
            final long longValue = lenientCellValue.longValue();
            assertEquals(1L, longValue);
        });
    }

}
