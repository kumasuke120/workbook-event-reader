package com.github.kumasuke120.excel;

import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Unmodifiable;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.TemporalAccessor;
import java.util.Collections;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;

/**
 * A collection of all {@link DateTimeFormatter} could be used during reading the workbook, providing
 * a set of methods to manipulate them
 */
@ApiStatus.Internal
class WorkbookDateTimeFormatters {

    // region predefined workbook formatters
    public static final DateTimeFormatter CN_DATE = DateTimeFormatter.ofPattern("yyyy'年'M'月'd'日'");
    public static final DateTimeFormatter CN_SLASH_DATE_TIME_AM_PM =
            DateTimeFormatter.ofPattern("yyyy/M/d h:mm a", Locale.US);
    public static final DateTimeFormatter CN_SLASH_DATE_TIME = DateTimeFormatter.ofPattern("yyyy/M/d[ H:mm]");

    public static final DateTimeFormatter EN_SLASH_DATE_TIME_AM_PM =
            DateTimeFormatter.ofPattern("M/d/yyyy h:mm a", Locale.US);
    public static final DateTimeFormatter EN_SLASH_DATE_TIME = DateTimeFormatter.ofPattern("M/d/yyyy[ H:mm]");
    public static final DateTimeFormatter EN_SLASH_DATE = DateTimeFormatter.ofPattern("MM/dd/yy H:m");
    public static final DateTimeFormatter EN_TIME_AM_PM =
            DateTimeFormatter.ofPattern("h:mm[:ss] a", Locale.US);
    // endregion

    /**
     * The default <code>DateTimeFormatter</code>s that will be used when converting <code>String</code>
     * to <code>LocalTime</code>, <code>LocalDate</code>, <code>LocalDateTime</code>
     */
    static final Set<DateTimeFormatter> DEFAULT_DATE_TIME_FORMATTERS = getPredefinedDateTimeFormatters();

    private WorkbookDateTimeFormatters() {
        throw new UnsupportedOperationException();
    }

    @NotNull
    static TemporalAccessor parseTemporalAccessor(@NotNull String value,
                                                  @NotNull Iterable<DateTimeFormatter> formatters) {
        RuntimeException previousEx = null;

        for (DateTimeFormatter formatter : formatters) {
            if (formatter == null) {
                if (previousEx == null) {
                    previousEx = new NullPointerException();
                } else {
                    previousEx.addSuppressed(new NullPointerException());
                }
                continue;
            }

            try {
                return formatter.parse(value);
            } catch (DateTimeParseException e) {
                if (previousEx != null) {
                    previousEx.addSuppressed(e);
                }
                previousEx = e;
            }
        }

        assert previousEx != null;
        throw new CellValueCastException(previousEx);
    }

    @Unmodifiable
    @NotNull
    static Set<DateTimeFormatter> getPredefinedDateTimeFormatters() {
        final Set<DateTimeFormatter> jdkFormatters = getPredefinedDateTimeFormatters(DateTimeFormatter.class);
        final Set<DateTimeFormatter> result = new HashSet<>(jdkFormatters);

        final Set<DateTimeFormatter> customizedFormatters =
                getPredefinedDateTimeFormatters(WorkbookDateTimeFormatters.class);
        result.addAll(customizedFormatters);

        return Collections.unmodifiableSet(result);
    }

    @NotNull
    private static Set<DateTimeFormatter> getPredefinedDateTimeFormatters(@NotNull Class<?> clazz) {
        final Set<DateTimeFormatter> formatters = new HashSet<>();
        final Field[] fields = clazz.getFields();
        for (Field field : fields) {
            final int modifiers = field.getModifiers();

            // public static final DateTimeFormatter ...
            if (Modifier.isStatic(modifiers) && Modifier.isFinal(modifiers) &&
                    DateTimeFormatter.class.equals(field.getType())) {
                try {
                    final DateTimeFormatter f = (DateTimeFormatter) field.get(null);
                    formatters.add(f);
                } catch (IllegalAccessException e) {
                    throw new AssertionError(e);
                }
            }
        }
        return formatters;
    }

}
