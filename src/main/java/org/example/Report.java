package org.example;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

public class Report {
    private static final Logger logger = LogManager.getLogger(Report.class);

    static final CellReference teacherSignatureCellRef = new CellReference("D20");
    static final CellReference bossSignatureCellRef = new CellReference("D21");
    static final CellReference reportYearCellRef = new CellReference("D5");
    static final CellReference reportMonthCellRef = new CellReference("D6");
    static final CellReference reportTotalDurationCellRef = new CellReference("D18");


    static int leftTopTableRow = 15;
    static int columnCount = 15;
    static Sheet reportSheet;
    static List<ReportRow> reportRowList;

    static String teacherSignature;
    static String bossSignature;
    static String totalDuration;

    public Report(Sheet reportSheet, List<ReportRow> reportRowList) {
        Report.reportSheet = reportSheet;
        Report.reportRowList = reportRowList;

        // вытаскиваем из шаблона и временно удаляем
        teacherSignature = saveAndDeleteCell(teacherSignatureCellRef);
        bossSignature = saveAndDeleteCell(bossSignatureCellRef);
        totalDuration = saveAndDeleteCell(reportTotalDurationCellRef);
    }

    private static Font fillFontStyle() {
        Font font = reportSheet.getWorkbook().createFont();
        // Настройки шрифта
        font.setFontName("Times New Roman");  // Times New Roman
        font.setFontHeightInPoints((short) 14); // Размер 14 pt
        font.setColor(IndexedColors.BLACK.getIndex()); // Черный цвет

        return font;
    }

    private static void fillBorderAtRows(int columnStart, int columnEnd, int rowIndex) {
        CellStyle cellBorderStyle = reportSheet.getWorkbook().createCellStyle();

        cellBorderStyle.setFont(fillFontStyle());  // Применяем шрифт к стилю
        // Обводка (черные линии)
        cellBorderStyle.setBorderTop(BorderStyle.THIN);
        cellBorderStyle.setBorderBottom(BorderStyle.THIN);
        cellBorderStyle.setBorderLeft(BorderStyle.THIN);
        cellBorderStyle.setBorderRight(BorderStyle.THIN);

        // Цвет границ (по умолчанию черный)
        cellBorderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellBorderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellBorderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellBorderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());

        // columnReportTable.setCellStyle(cellStyle);

        Row row = reportSheet.getRow(rowIndex);

        for (int columnIndex = columnStart; columnIndex < columnEnd; columnIndex++) {
            Cell cell = row.getCell(columnIndex);
            if (cell == null) {
                cell = row.createCell(columnIndex);
            }
            cell.setCellStyle(cellBorderStyle);
        }
    }

    // вставляем нижнюю часть отчёта
    private static void printFooterReport() {
        fillCell(new CellReference(Report.reportRowList.size() - 1 + teacherSignatureCellRef.getRow(), teacherSignatureCellRef.getCol()), teacherSignature);
        fillCell(new CellReference(Report.reportRowList.size() - 1 + bossSignatureCellRef.getRow(), bossSignatureCellRef.getCol()), bossSignature);
        fillCell(new CellReference(Report.reportRowList.size() - 1 + reportTotalDurationCellRef.getRow(), reportTotalDurationCellRef.getCol()), totalDuration);

        fillBorderAtRows(0, 15, Report.reportRowList.size() - 1 + reportTotalDurationCellRef.getRow());
    }

    private String saveAndDeleteCell(CellReference cellReference) {
        Row row = reportSheet.getRow(cellReference.getRow());
        Cell cell = row.getCell(cellReference.getCol());

        String saveVar = cell.getStringCellValue();
        cell.setCellValue("");

        return saveVar;
    }

    private static void fillCell(CellReference cellReference, String value) {
        Row row = reportSheet.getRow(cellReference.getRow());
        Cell cell = row.getCell(cellReference.getCol());
        CellStyle cellBorderStyle = reportSheet.getWorkbook().createCellStyle();
        cellBorderStyle.setFont(fillFontStyle());

        if (cell == null) {
            logger.error("Ячейка {} равна null", cellReference.formatAsString());
            cell = row.createCell(cellReference.getCol());

        }
        cell.setCellStyle(cellBorderStyle);
        cell.setCellValue(value);
    }

    private Map<Integer, Function<ReportRow, String>> getColumnValueMap() {
        Map<Integer, Function<ReportRow, String>> columnValueMap = new HashMap<>();
        columnValueMap.put(1, ReportRow::getReportLessonDayString);
        columnValueMap.put(2, ReportRow::getReportLessonTime);
        columnValueMap.put(3, ReportRow::getReportLessonName);
        columnValueMap.put(6, ReportRow::getReportGroupName);
        columnValueMap.put(7, row -> "30"); // Фиксированные значения
        columnValueMap.put(8, row -> "2");

        return columnValueMap;
    }

    public void printReportRows() {
        // Маппинг: номер колонки -> лямбда, возвращающая значение
        Map<Integer, Function<ReportRow, String>> columnValueMap = new HashMap<>();
        columnValueMap.put(0, ReportRow::getReportLessonDayString);
        columnValueMap.put(1, ReportRow::getReportLessonTime);
        columnValueMap.put(2, ReportRow::getReportLessonName);
        columnValueMap.put(5, ReportRow::getReportGroupName);
        columnValueMap.put(6, row -> "30"); // Фиксированные значения
        columnValueMap.put(7, row -> "2");

        for (ReportRow reportRow : reportRowList) {
            // ReportRow.printReportRow(reportRow);

            Row rowReportTable = reportSheet.getRow(leftTopTableRow);
            if (rowReportTable == null) {
                rowReportTable = reportSheet.createRow(leftTopTableRow);
            }

            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                // Создаем стиль для всех ячеек
                CellStyle cellStyle = createCellStyle();
                Cell columnReportTable = rowReportTable.getCell(columnIndex);
                if (columnReportTable == null) {
                    break;
                }
                columnReportTable.setCellStyle(cellStyle);

                if (columnValueMap.containsKey(columnIndex)) {
                    if (columnIndex == 6 || columnIndex == 7) {
                        columnReportTable.setCellValue(Integer.parseInt(columnValueMap.get(columnIndex).apply(reportRow)));
                    } else {
                        columnReportTable.setCellValue(columnValueMap.get(columnIndex).apply(reportRow));
                    }

                }
            }

            leftTopTableRow++;
        }
        printFooterReport();
    }

    // 📌 Метод для создания стиля с выравниванием и обводкой
    private CellStyle createCellStyle() {
        CellStyle style = reportSheet.getWorkbook().createCellStyle();

        // Выравнивание по центру
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // Обводка (черные линии)
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        // Цвет границ (по умолчанию черный)
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());

        style.setFont(fillFontStyle());  // Применяем шрифт к стилю

        return style;
    }
}
