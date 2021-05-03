package com.github.kumasuke120.excel;

import org.junit.jupiter.api.Test;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import static org.junit.jupiter.api.Assertions.assertEquals;

class WorkbookDateTimeFormattersTest {

    @Test
    void cnPredefinedFormatters() {
        final LocalDate date1 = WorkbookDateTimeFormatters.CN_DATE.parse("2021年4月12日", LocalDate::from);
        assertEquals(LocalDate.of(2021, 4, 12), date1);

        final LocalDateTime dateTime1 = WorkbookDateTimeFormatters.CN_SLASH_DATE_TIME_AM_PM
                .parse("2018/11/9 8:00 PM", LocalDateTime::from);
        assertEquals(LocalDateTime.of(2018, 11, 9, 20, 0, 0), dateTime1);

        final LocalDateTime dateTime2 = WorkbookDateTimeFormatters.CN_SLASH_DATE_TIME
                .parse("2018/11/9 20:00", LocalDateTime::from);
        assertEquals(LocalDateTime.of(2018, 11, 9, 20, 0, 0), dateTime2);

        final LocalDate date2 = WorkbookDateTimeFormatters.CN_SLASH_DATE_TIME
                .parse("2018/11/9", LocalDate::from);
        assertEquals(LocalDate.of(2018, 11, 9), date2);
    }

    @Test
    void enPredefinedFormatters() {
        final LocalDateTime dateTime3 = WorkbookDateTimeFormatters.EN_SLASH_DATE_TIME_AM_PM
                .parse("11/9/2018 8:00 PM", LocalDateTime::from);
        assertEquals(LocalDateTime.of(2018, 11, 9, 20, 0, 0), dateTime3);

        final LocalDateTime dateTime4 = WorkbookDateTimeFormatters.EN_SLASH_DATE_TIME
                .parse("11/9/2018 20:00", LocalDateTime::from);
        assertEquals(LocalDateTime.of(2018, 11, 9, 20, 0, 0), dateTime4);

        final LocalDate date3 = WorkbookDateTimeFormatters.EN_SLASH_DATE_TIME
                .parse("11/9/2018", LocalDate::from);
        assertEquals(LocalDate.of(2018, 11, 9), date3);

        final LocalDateTime dateTime5 = WorkbookDateTimeFormatters.EN_SLASH_DATE
                .parse("11/09/18 20:00", LocalDateTime::from);
        assertEquals(LocalDateTime.of(2018, 11, 9, 20, 0, 0), dateTime5);

        final LocalTime time1 = WorkbookDateTimeFormatters.EN_TIME_AM_PM
                .parse("8:00:08 PM", LocalTime::from);
        assertEquals(LocalTime.of(20, 0, 8), time1);

        final LocalTime time2 = WorkbookDateTimeFormatters.EN_TIME_AM_PM
                .parse("8:00 PM", LocalTime::from);
        assertEquals(LocalTime.of(20, 0, 0), time2);
    }

}