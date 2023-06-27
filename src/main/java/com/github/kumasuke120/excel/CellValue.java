package com.github.kumasuke120.excel;

import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.function.Function;

/**
 * A value class which encapsulates all possible values that the {@link WorkbookEventReader} may return and which
 * provides convenient ways to convert between them
 */
public interface CellValue {

    /**
     * Returns a <code>CellValue</code> based on the given value.
     *
     * @param originalValue the given value
     * @return an instance of <code>CellValue</code>
     */
    @NotNull
    static CellValue newInstance(@Nullable Object originalValue) {
        return LenientCellValue.newInstance(originalValue);
    }

    /**
     * Returns type of the original value. It will return <code>null</code> if the original value is <code>null</code>.
     *
     * @return the type of the original value if possible, otherwise a {@link NullPointerException} will be thrown
     * @throws NullPointerException the original value is null
     */
    @NotNull
    Class<?> originalType();

    /**
     * Returns the original value of this {@link CellValue}.<br>
     * The original value will be one of the following types:
     * <ul>
     * <li><code>null</code></li>
     * <li>{@link Boolean}</li>
     * <li>{@link Integer}</li>
     * <li>{@link Long}</li>
     * <li>{@link Double}</li>
     * <li>{@link String}</li>
     * <li>{@link LocalTime}</li>
     * <li>{@link LocalDate}</li>
     * <li>{@link LocalDateTime}</li>
     * </ul>
     *
     * @return original value
     */
    @Nullable
    Object originalValue();

    /**
     * Checks if the original value is <code>null</code>.
     *
     * @return <code>true</code> if the original value is <code>null</code>, otherwise <code>false</code>
     */
    boolean isNull();

    /**
     * Converts the original value to its <code>boolean</code> counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type <code>boolean</code>, it will
     * be returned directly.
     *
     * @return <code>boolean</code> version of the original value
     * @throws CellValueCastException cannot convert the original value to <code>boolean</code> type
     */
    boolean booleanValue();

    /**
     * Converts the original value to its <code>int</code> counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type <code>int</code>, it will
     * be returned directly.
     *
     * @return <code>int</code> version of the original value
     * @throws CellValueCastException cannot convert the original value to <code>int</code> type
     */
    int intValue();

    /**
     * Converts the original value to its <code>long</code> counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type <code>long</code>, it will
     * be returned directly.
     *
     * @return <code>long</code> version of the original value
     * @throws CellValueCastException cannot convert the original value to <code>long</code> type
     */
    long longValue();

    /**
     * Converts the original value to its <code>double</code> counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type <code>double</code>, it will
     * be returned directly.
     *
     * @return <code>double</code> version of the original value
     * @throws CellValueCastException cannot convert the original value to <code>double</code> type
     */
    double doubleValue();

    /**
     * Converts the original value to its {@link String} counterpart.<br>
     * The conversion only happens when it's possible. If the original value is of type {@link String}, it will
     * be returned directly. Otherwise, the original value will be converted to {@link String} via
     * {@link String#valueOf(Object)}.
     *
     * @return {@link String} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     */
    @NotNull
    String stringValue();

    /**
     * Converts the original value to its {@link LocalTime} counterpart.<br>
     * It will use all preset {@link DateTimeFormatter} static fields that is of type {@link DateTimeFormatter}
     * if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalTime}, it will
     * be returned directly.
     *
     * @return {@link LocalTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     */
    @NotNull
    LocalTime localTimeValue();

    /**
     * Converts the original value to its {@link LocalTime} counterpart.<br>
     * The given formatter will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalTime}, it will
     * be returned directly.
     *
     * @param formatter {@link DateTimeFormatter} instance to help converting {@link String} to {@link LocalTime}
     * @return {@link LocalTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatter is <code>null</code>
     */
    @NotNull
    LocalTime localTimeValue(@NotNull(exception = NullPointerException.class) DateTimeFormatter formatter);

    /**
     * Converts the original value to its {@link LocalTime} counterpart.<br>
     * The given formatters will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalTime}, it will
     * be returned directly.
     *
     * @param formatters {@link DateTimeFormatter} instances to help converting {@link String} to {@link LocalTime}
     * @return {@link LocalTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatters is <code>null</code>
     */
    @NotNull
    LocalTime localTimeValue(@NotNull(exception = NullPointerException.class) Iterable<DateTimeFormatter> formatters);

    /**
     * Converts the original value to its {@link LocalDate} counterpart.<br>
     * It will use all preset {@link DateTimeFormatter} static fields that is of type {@link DateTimeFormatter}
     * if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDate}, it will
     * be returned directly.
     *
     * @return {@link LocalDate} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     */
    @NotNull
    LocalDate localDateValue();

    /**
     * Converts the original value to its {@link LocalDate} counterpart.<br>
     * The given formatter will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDate}, it will
     * be returned directly.
     *
     * @param formatter {@link DateTimeFormatter} instance to help converting {@link String} to {@link LocalDate}
     * @return {@link LocalDate} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatter is <code>null</code>
     */
    @NotNull
    LocalDate localDateValue(@NotNull(exception = NullPointerException.class) DateTimeFormatter formatter);

    /**
     * Converts the original value to its {@link LocalDate} counterpart.<br>
     * The given formatters will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDate}, it will
     * be returned directly.
     *
     * @param formatters {@link DateTimeFormatter} instances to help converting {@link String} to {@link LocalDate}
     * @return {@link LocalDate} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatters is <code>null</code>
     */
    @NotNull
    LocalDate localDateValue(@NotNull(exception = NullPointerException.class) Iterable<DateTimeFormatter> formatters);

    /**
     * Converts the original value to its {@link LocalDateTime} counterpart.<br>
     * It will use all preset {@link DateTimeFormatter} static fields that is of type {@link DateTimeFormatter}
     * if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDateTime}, it will
     * be returned directly.
     *
     * @return {@link LocalDateTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     */
    @NotNull
    LocalDateTime localDateTimeValue();

    /**
     * Converts the original value to its {@link LocalDateTime} counterpart.<br>
     * The given formatter will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDateTime}, it will
     * be returned directly.
     *
     * @param formatter {@link DateTimeFormatter} instance to help converting {@link String} to {@link LocalDateTime}
     * @return {@link LocalDateTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatter is <code>null</code>
     */
    @NotNull
    LocalDateTime localDateTimeValue(@NotNull(exception = NullPointerException.class) DateTimeFormatter formatter);

    /**
     * Converts the original value to its {@link LocalDateTime} counterpart.<br>
     * The given formatters will be used if the original value is of type {@link String} to help parsing.
     * The conversion only happens when it's possible. If the original value is of type {@link LocalDateTime}, it will
     * be returned directly.
     *
     * @param formatters {@link DateTimeFormatter} instances to help converting {@link String} to {@link LocalDateTime}
     * @return {@link LocalDateTime} version of the original value
     * @throws CellValueCastException cannot convert the original value to {@link String} type
     * @throws NullPointerException   the given formatters is <code>null</code>
     */
    @NotNull
    LocalDateTime localDateTimeValue(@NotNull(exception = NullPointerException.class)
                                     Iterable<DateTimeFormatter> formatters);

    /**
     * Uses the given function to map the {@link #originalValue()} to a new value, and returns a new
     * {@link CellValue} based on the new value.<br>
     * The return value of <code>mappingFunction</code> must be one of the types listed on the javadoc of
     * {@link #originalValue()}.
     *
     * @param mappingFunction value mapping function
     * @return {@link CellValue} based on the mapped value
     * @throws CellValueCastException cannot perform mapping or the mapped value is not in the allow type
     */
    @NotNull
    @Contract(pure = true)
    CellValue mapOriginalValue(@NotNull(exception = NullPointerException.class)
                               Function<Object, Object> mappingFunction);

    /**
     * Trims the original value if its type is of {@link String}.<br>
     * A new instance of {@link CellValue} will be created if there is any change to this
     * {@link CellValue} instance.
     *
     * @return {@link CellValue} based on the trimmed value
     */
    @NotNull
    @Contract(pure = true)
    CellValue trim();

}
