package org.example;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ReportRow {
    private static final Logger logger = LogManager.getLogger(ReportRow.class);
    private final Date reportLessonDay;
    private final String reportLessonTime;
    private String reportLessonName;
    private String reportGroupName;
    private String reportLessonDayString;

    public ReportRow(Date reportLessonDay, String reportLessonTime, String reportLessonName, String reportGroupName) {
        this.reportLessonDay = reportLessonDay;
        this.reportLessonTime = reportLessonTime;
        this.reportLessonName = reportLessonName;
        this.reportGroupName = reportGroupName;
    }

    public Date getReportLessonDay() {
        return reportLessonDay;
    }

    public String getReportLessonTime() {
        return reportLessonTime;
    }

    public String getReportLessonName() {
        return reportLessonName;
    }

    public String getReportGroupName() {
        return JSONTranslator.translate(reportGroupName, reportGroupName);
    }

    public String getReportLessonDayString() {
        return reportLessonDayString;
    }

    public static void sort(List<ReportRow> reportRowList) {
        reportRowList.sort(Comparator.comparing(ReportRow::getReportLessonDay).thenComparing(ReportRow::getReportLessonTime));

        for (ReportRow reportRow : reportRowList) {
            SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");
            reportRow.reportLessonDayString = sdf.format(reportRow.reportLessonDay);
        }

        logger.debug("Отсортировано");
    }

    public static void printReportRows(List<ReportRow> reportRowList) {
        for (ReportRow reportRow : reportRowList) {
            printReportRow(reportRow);
        }
    }

    public static void printReportRow(ReportRow reportRow) {
        logger.info("{} {} {} {}", reportRow.reportLessonDayString, reportRow.reportLessonTime, reportRow.reportLessonName, reportRow.reportGroupName);
    }


    public static void mergeReportRows(List<ReportRow> reportRowList) {
        int i = 0;
        while (i < reportRowList.size() - 1) {
            ReportRow currentReportRow = reportRowList.get(i);
            ReportRow nextReportRow = reportRowList.get(i + 1);
            if (currentReportRow.equals(nextReportRow)) {
                currentReportRow.reportGroupName = ReportRow.mergeGroupName(currentReportRow.reportGroupName, nextReportRow.reportGroupName);
                currentReportRow.reportLessonName = currentReportRow.reportLessonName.trim();
                reportRowList.remove(i + 1);
            } else {
                i++;
            }
        }

        logger.debug("Списки оформили слияние");
    }

    public boolean equals(ReportRow reportRowList) {
        return this.reportLessonDay.equals(reportRowList.reportLessonDay)
                && this.reportLessonName.equals(reportRowList.reportLessonName)
                && this.reportLessonTime.equals(reportRowList.reportLessonTime);
    }

    public static String mergeGroupName(String firstGroupName, String secondGroupName) {
        // Регулярное выражение для поиска двузначных чисел
        Pattern pattern = Pattern.compile("\\d{2}");

        Matcher firstMatcher = pattern.matcher(firstGroupName);
        Matcher secondMatcher = pattern.matcher(secondGroupName);

        if (firstMatcher.find() && secondMatcher.find()) {
            int firstNumber = Integer.parseInt(firstMatcher.group());
            int secondNumber = Integer.parseInt(secondMatcher.group());

            if (firstNumber > secondNumber) {
                return String.format("ПФ-%d-%d", secondNumber, firstNumber);
            } else {
                return String.format("ПФ-%d-%d", firstNumber, secondNumber);
            }
        } else {
            throw new RuntimeException("Чето по пизде пошло");
        }
    }

    public static List<ReportRow> filterReportRows(List<ReportRow> reportRowList) {

        List<ReportRow> newReportRowList = new ArrayList<>();

        for (ReportRow reportRow: reportRowList){
            if (TimeRange.checkDate(reportRow.getReportLessonDay())){
                newReportRowList.add(reportRow);
            }
        }

        printReportRows(newReportRowList);

        return newReportRowList;
    }
}

