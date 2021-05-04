package com.github.kumasuke120.excel;

import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DateUtil;

import java.time.*;
import java.util.Collections;
import java.util.Date;
import java.util.Map;
import java.util.TimeZone;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * An utility class that contains various methods for dealing with workbook
 */
class Util {

    private static final double MAX_EXCEL_DATE_EXCLUSIVE = 2958466;
    private static final Pattern cellReferencePattern = Pattern.compile("([A-Z]+)(\\d+)");

    private Util() {
        throw new UnsupportedOperationException();
    }

    /**
     * Tests if the given value is a whole number.
     *
     * @param value <code>double</code> value to be tested
     * @return <code>true</code> if the given value is a whole number, otherwise <code>false</code>
     */
    static boolean isAWholeNumber(double value) {
        return value % 1 == 0;
    }

    /**
     * Tests if the given value is a whole number that Java could represent with primitive type.
     *
     * @param value <code>String</code> value to be tested
     * @return <code>true</code> if the given value is a whole number, otherwise <code>false</code>
     */
    static boolean isAWholeNumber(String value) {
        if (value == null) return false;

        try {
            Long.parseLong(value);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * Tests if the given value is a decimal fraction that Java could represent with primitive type.
     *
     * @param value <code>String</code> value to be tested
     * @return <code>true</code> if the given value is a decimal fraction, otherwise <code>false</code>
     */
    static boolean isADecimalFraction(String value) {
        if (value == null) return false;

        try {
            Double.parseDouble(value);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    /**
     * Parses the <code>String</code> value to <code>int</code> silently.<br>
     * It will return <code>defaultValue</code> if the <code>String</code> value could not be parsed.
     *
     * @param value the <code>String</code> value to be parsed
     * @return parsed <code>int</code> value if parse succeeds, otherwise the <code>defaultValue</code>
     */
    @SuppressWarnings("SameParameterValue")
    static int toInt(String value, int defaultValue) {
        if (value == null) return defaultValue;

        try {
            return Integer.parseInt(value);
        } catch (NumberFormatException e) {
            return defaultValue;
        }
    }

    /**
     * Tests if the given value is a valid excel date.
     *
     * @param excelDateValue excel date value
     * @return code>true</code> if the given value is a valid excel date, otherwise <code>false</code>
     */
    static boolean isValidExcelDate(double excelDateValue) {
        return DateUtil.isValidExcelDate(excelDateValue) &&
                excelDateValue > 0 && excelDateValue < MAX_EXCEL_DATE_EXCLUSIVE;
    }

    /**
     * Converts the given excel date value to @{@link LocalTime}, {@link LocalDateTime} or
     * {@link LocalDate} accordingly.
     *
     * @param excelDateValue   excel date value
     * @param use1904Windowing <code>true</code> if date uses 1904 windowing,
     *                         or <code>false</code> if using 1900 date windowing.
     * @return converted @{@link LocalTime}, {@link LocalDateTime} or {@link LocalDate}
     */
    static Object toJsr310DateOrTime(double excelDateValue, boolean use1904Windowing) {
        if (isValidExcelDate(excelDateValue)) {
            final Date date = DateUtil.getJavaDate(excelDateValue, use1904Windowing,
                                                   TimeZone.getTimeZone("UTC"));
            final LocalDateTime localDateTime = toLocalDateTimeOffsetByUTC(date);

            if (Util.isAWholeNumber(excelDateValue)) { // date only
                return localDateTime.toLocalDate();
            } else if (excelDateValue < 1) { // time only
                return localDateTime.toLocalTime();
            } else { // date with time
                return localDateTime;
            }
        } else {
            return null;
        }
    }

    private static LocalDateTime toLocalDateTimeOffsetByUTC(Date date) {
        final Instant instant = date.toInstant();
        return LocalDateTime.ofInstant(instant, ZoneId.of("UTC"));
    }

    /**
     * Converts given value to relative types.<br>
     * For instance, if a <code>double</code> could be treated as a <code>int</code>, it will be converted
     * to its <code>int</code> counterpart.
     *
     * @param value value to be converted
     * @return value converted accordingly
     */
    static Object toRelativeType(Object value) {
        if (value instanceof Double) {
            final double doubleValue = (double) value;
            if (Util.isAWholeNumber(doubleValue)) {
                if (doubleValue > Integer.MAX_VALUE || doubleValue < Integer.MIN_VALUE) {
                    return (long) doubleValue;
                } else {
                    return (int) doubleValue;
                }
            }
        } else if (value instanceof Long) {
            final long longValue = (long) value;
            if (longValue >= Integer.MIN_VALUE && longValue <= Integer.MAX_VALUE) {
                return (int) longValue;
            }
        }

        return value;
    }

    /**
     * Converts excel cell reference (e.g. A1) to a {@link Map.Entry} containing row number (starts with 0)
     * and column number (starts with 0).<br>
     * It will return <code>null</code> if either of them cannot be parse correctly.
     *
     * @param cellReference excel cell reference
     * @return row number and column number, both starts with 0 or <code>null</code>
     */
    static Map.Entry<Integer, Integer> cellReferenceToRowAndColumn(String cellReference) {
        assert cellReference != null;

        final Matcher cellRefMatcher = cellReferencePattern.matcher(cellReference);
        if (cellRefMatcher.matches()) {
            final String rawColumn = cellRefMatcher.group(1);
            final String rawRow = cellRefMatcher.group(2);

            final int rowNum = Integer.parseInt(rawRow) - 1;
            final int columnNum = columnNameToInt(rawColumn);

            return Collections.singletonMap(rowNum, columnNum)
                    .entrySet()
                    .iterator()
                    .next();
        } else {
            return null;
        }
    }

    // converts 'A' to 0, 'B' to 1, ..., 'AA' to 26, and etc.
    private static int columnNameToInt(String columnName) {
        int index = 0;

        for (int i = 0; i < columnName.length(); i++) {
            final char c = columnName.charAt(i);
            if (i == columnName.length() - 1) {
                index = index * 26 + (c - 'A');
            } else {
                index = index * 26 + (c - 'A' + 1);
            }
        }

        return index;
    }

    /**
     * Tests if the given index of format or format string stands for a text format.
     *
     * @param formatIndex  index of format
     * @param formatString format string
     * @return <code>true</code> if the given arguments stand for a text format, otherwise <code>false</code>
     */
    static boolean isATextFormat(int formatIndex, String formatString) {
        return formatIndex == 0x31 ||
                (formatString != null && BuiltinFormats.getBuiltinFormat(formatString) == 0x31);
    }

}
