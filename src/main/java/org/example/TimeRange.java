package org.example;

import java.time.LocalDate;
import java.time.ZoneId;
import java.time.temporal.TemporalAdjusters;
import java.util.Date;


public class TimeRange {
    private static LocalDate startDate;
    private static LocalDate endDate;

    public TimeRange(int month) {
         startDate = LocalDate.of(2025, month, 1);
         endDate = startDate.with(TemporalAdjusters.lastDayOfMonth());
    }

    public static boolean checkDate(Date dateLesson) {

        LocalDate dateLessonLocalDate = DatetoLocalDate(dateLesson);

        return !(dateLessonLocalDate).isAfter((endDate)) && !(dateLessonLocalDate).isBefore((startDate));
    }

    private static LocalDate DatetoLocalDate(Date date){
        return date.toInstant()
                .atZone(ZoneId.systemDefault())
                .toLocalDate();
    }
}
