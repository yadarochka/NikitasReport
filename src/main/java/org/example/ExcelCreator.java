package org.example;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelCreator {
    private static final Logger logger = LogManager.getLogger(ExcelCreator.class);

    public static void createFile(List<ReportRow> reportRowList) throws IOException, InterruptedException {
        // Указываем путь к Excel файлу
        String filePath = ConfigLoader.getTemplateFilePath();

        // Считываем файл
        FileInputStream file = new FileInputStream(filePath);

        // Создаем копию на основе шаблона
        Workbook workbook = WorkbookFactory.create(file);

        // Ищем книгу
        Sheet sheet = workbook.getSheet("отчет");


        logger.info(sheet.getSheetName());

        workbook.setSheetName(workbook.getSheetIndex(sheet), "Щербаков отчёт февраль");

        logger.info(sheet.getSheetName());


        // Если книги нет, завершаем работу
        if (sheet == null){
            throw new RuntimeException("Файл-шаблон не найден");
        }

        Report report = new Report(sheet, reportRowList);

        report.printReportRows();

        FileOutputStream outputStream = new FileOutputStream(ConfigLoader.getOutputFilePath());
        workbook.write(outputStream);

        workbook.close();
        file.close();
    }
}
