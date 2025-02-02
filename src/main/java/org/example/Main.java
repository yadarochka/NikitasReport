package org.example;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.*;
import java.text.ParseException;
import java.util.List;
import java.util.Scanner;


public class Main {
    private static final Logger logger = LogManager.getLogger(Main.class);

    public static void main(String[] args) throws IOException, InterruptedException, ParseException {
        Scanner scanner = new Scanner(System.in);
        int month = scanner.nextInt();
        scanner.close();

        if (month < 1 && month > 12) {
            return;
        }

        TimeRange timerange = new TimeRange(month);

        FileHelper.downloadFile();
        List<ReportRow> reportRowList = ReportRow.filterReportRows(ExcelReader.readFile());
        ReportRow.printReportRows(reportRowList);
        ReportRow.sort(reportRowList);
        ReportRow.mergeReportRows(reportRowList);
        ReportRow.printReportRows(reportRowList);
        ExcelCreator.createFile(reportRowList);
        FileHelper.deleteFile(FileHelper.FILE_TYPE.INPUT);
        FileHelper.openOutputFile();
    }
}
