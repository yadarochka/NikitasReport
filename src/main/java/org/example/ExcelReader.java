package org.example;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.*;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelReader {

    private static final Logger logger = LogManager.getLogger(ExcelReader.class);
    private static CellReference topLeftCornerCell = new CellReference("B3");
    private static final int rowCount = 5;
    private static final int columnCount = 6;
    private static final int groupStep = 7;
    private static final int groupCount = 28;
    private static int numberWeekStart = 18;
    private static int numberWeekEnd = 28;

    public static List<ReportRow> readFile() throws IOException {
        logger.debug("Чтение файла excel...");
        String filePath = ConfigLoader.getInputFilePath(); // Путь к файлу Excel
        List<ReportRow> reportRowList = new ArrayList<>();

        // Указываем путь к Excel файлу
        FileInputStream file = new FileInputStream(filePath);

        // Получаем рабочую книгу
        Workbook workbook = WorkbookFactory.create(file);

        int currentWeek = numberWeekStart;

        while(currentWeek < numberWeekEnd) {
            logger.info(currentWeek);
            String sheetName = String.valueOf(currentWeek);
            Sheet sheet = findSheetByPrefix(workbook, sheetName);

            if (sheet != null) {
                logger.info("Найден лист с именем: {}", sheet.getSheetName());
            } else {
                logger.error("Лист с таким именем не найден.");
                if (reportRowList.size() == 0) {
                    throw new RuntimeException("Занятия не нашлись =(");
                }
                break;
            }

            findRowList(sheet, reportRowList);

            currentWeek++;

        }
        workbook.close();
        file.close();
        return reportRowList;
    }

    private static Sheet findSheetByPrefix(Workbook workbook, String prefix) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet.getSheetName().startsWith(prefix)) {
                return sheet; // Возвращаем первый найденный лист с подходящим именем
            }
        }
        return null; // Если лист с таким именем не найден
    }

    // получаем название группы по верхней левой ячейке её блока
    private static String getGroupName(Sheet sheet, CellReference topLeftCornerCell) {
        // сдвигаемся для получения названия группы
        int groupNameRow = topLeftCornerCell.getRow() + 1;
        int groupNameCell = topLeftCornerCell.getCol();

        Row row = sheet.getRow(groupNameRow);

        if (row == null) {
            logger.fatal("Название группы получить невозможно. Ячейка:{}", new CellReference(groupNameRow, groupNameCell).formatAsString());
            return "";
        }

        Cell cell = row.getCell(groupNameCell);

        if (cell.getStringCellValue().trim().split(" ").length == 1) {
            return cell.getStringCellValue().trim();
        }

        if (cell.getStringCellValue().trim().split(" ").length > 1) {
            for (String str : cell.getStringCellValue().trim().split(" ")){
                if (str.startsWith("ПФ")){
                    return str.trim();
                }
            }
        }

        throw new RuntimeException("Пизда рулю");
    }

    // получаем дату занятия по верхней левой ячейке её блока и столбцу с фамилией
    private static Date getLessonDate(Sheet sheet, CellReference topLeftCornerCell, int lessonDateCell) {
        int lessonDateRow = topLeftCornerCell.getRow();

        Row row = sheet.getRow(lessonDateRow);

        CellReference currentCell = new CellReference(lessonDateRow, lessonDateCell);

        //logger.info("Ищем дату в ячейке {}",currentCell.formatAsString());

        if (row == null) {
            logger.fatal("Дату занятия получить невозможно. Ячейка:{}", new CellReference(lessonDateRow, lessonDateCell).formatAsString());
            throw new RuntimeException("Пизда рулю");
        }
        try {
            Cell cell = row.getCell(lessonDateCell);
            return cell.getDateCellValue();
        } catch (Exception e) {
            logger.fatal("Дату занятия получить невозможно. Ячейка:{}", new CellReference(lessonDateRow, lessonDateCell).formatAsString());
            throw new RuntimeException("Пизда рулю");
        }


    }

    private static String getLessonTime(Sheet sheet, CellReference topLeftCornerCell, int lessonTimeRow) {
        int lessonTimeCell = topLeftCornerCell.getCol();

        Row row = sheet.getRow(lessonTimeRow);

        CellReference currentCell = new CellReference(lessonTimeRow, lessonTimeCell);

        //logger.info("Ищем время в ячейке {}",currentCell.formatAsString());

        if (row == null) {
            logger.fatal("Время занятия получить невозможно. Ячейка:{}", currentCell.formatAsString());
            return "";
        }

        try {
            Cell cell = row.getCell(lessonTimeCell);
            //logger.info("Время: {}", cell.getStringCellValue());
            return cell.getStringCellValue().trim();
        } catch (Exception e) {
            logger.fatal("Время занятия получить невозможно. Ячейка:{}", currentCell.formatAsString());
            throw new RuntimeException(e);
        }
    }

    private static List<ReportRow> findRowList(Sheet sheet, List<ReportRow> reportRowList){
        int currentGroup = 0;

        topLeftCornerCell = new CellReference("B3");

        while (currentGroup < groupCount) {
            logger.info(currentGroup);
            String groupName = getGroupName(sheet, topLeftCornerCell);
            //logger.debug("{} Группа: {}", currentGroup + 1, groupName);
            // сдвигаемся для получения ячеек с предметами
            CellReference lessonCellStart = new CellReference((topLeftCornerCell.getRow()) + 2, topLeftCornerCell.getCol() + 1);

            for (int currentRowIndex = lessonCellStart.getRow(); currentRowIndex < lessonCellStart.getRow() + rowCount; currentRowIndex++) {
                for (int currentCellIndex = lessonCellStart.getCol(); currentCellIndex < lessonCellStart.getCol() + columnCount; currentCellIndex++) {
                    Row currentRow = sheet.getRow(currentRowIndex);
                    if (currentRow == null) {
                        logger.error("row = null");
                        break;
                    }
                    Cell currentCell = currentRow.getCell(currentCellIndex);

                    CellReference currentCellRef = new CellReference(currentRowIndex, currentCellIndex);

                    //logger.info("Проверяем ячейку {}", currentCellRef.formatAsString());

                    if (currentCell == null || currentCell.getCellType() != CellType.STRING) {
                        //logger.error("Пусто. Строка:{}, Тип:{}", currentRowIndex, currentCell.getCellType());
                        continue;
                    }

                    if (currentCell.getStringCellValue().toLowerCase().contains("щербаков")) {
                        //logger.info("Предмет: {}", currentCell.getStringCellValue().trim().split("\\s{2,}")[0]);
                        String lessonName = getLessonName(currentCell.getStringCellValue());
                        Date lessonDate = getLessonDate(sheet, topLeftCornerCell, currentCellIndex);
                        String lessonTime = getLessonTime(sheet, topLeftCornerCell, currentRowIndex);
                        //logger.info("Занятие у группы {} по предмету \"{}\" {} {}", groupName, lessonName, lessonDate, lessonTime);

                        reportRowList.add(new ReportRow(lessonDate, lessonTime, lessonName, groupName));
                    }
                }
            }

            currentGroup++;
            topLeftCornerCell = new CellReference(topLeftCornerCell.getRow() + groupStep, topLeftCornerCell.getCol());
        }

        return reportRowList;
    }

    private static String getLessonName(String str) {
        String regex = "(История России|Обществознание)";

        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(str);

        if (matcher.find()) {
            return matcher.group(); // Возвращаем найденное название предмета
        }
        else {
            throw new RuntimeException("В названии нет предмета");
        }
    }
}