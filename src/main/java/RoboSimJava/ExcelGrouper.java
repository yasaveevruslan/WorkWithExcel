package RoboSimJava;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Pattern;

public class ExcelGrouper {

    // ==================== ВСПОМОГАТЕЛЬНЫЕ КЛАССЫ ====================

    public static class DataRow {
        private String groupName;
        private final Object[] allData;
        private final double value1;
        private final double value2;
        private Integer year;
        private Date fullDate;
        private String inn;
        private String operationName;

        public DataRow(String groupName, Object[] allData, double value1, double value2) {
            this.groupName = groupName;
            this.allData = allData;
            this.value1 = value1;
            this.value2 = value2;
            this.inn = "";
            this.operationName = "";
            this.year = null;
            this.fullDate = null;
        }

        public String getGroupName() { return groupName; }
        public Object[] getAllData() { return allData; }
        public double getValue1() { return value1; }
        public double getValue2() { return value2; }
        public Integer getYear() { return year; }
        public Date getFullDate() { return fullDate; }
        public String getInn() { return inn; }
        public String getOperationName() { return operationName; }

        public void setGroupName(String groupName) { this.groupName = groupName; }
        public void setInn(String inn) { this.inn = inn; }
        public void setOperationName(String operationName) { this.operationName = operationName; }
        public void setYear(Integer year) { this.year = year; }

        public void setDate(Date date) {
            if (date != null) {
                this.fullDate = date;
                Calendar cal = Calendar.getInstance();
                cal.setTime(date);
                this.year = cal.get(Calendar.YEAR);
            }
        }

        public String getDisplayName() {
            if (inn != null && !inn.isEmpty()) {
                return groupName + " " + inn;
            }
            return groupName;
        }
    }

    public static class YearGroup {
        private final int year;
        private final List<DataRow> rows = new ArrayList<>();
        private double totalValue;

        public YearGroup(int year) {
            this.year = year;
            this.totalValue = 0;
        }

        public void addRow(DataRow row, double value) {
            rows.add(row);
            totalValue += value;
        }

        public int getYear() { return year; }
        public List<DataRow> getRows() { return rows; }
        public double getTotalValue() { return totalValue; }
        public int getRowCount() { return rows.size(); }

        public void sortRowsByDate() {
            rows.sort(Comparator.comparing(DataRow::getFullDate, Comparator.nullsLast(Comparator.naturalOrder())));
        }
    }

    public static class DataGroup {
        private final String name;
        private String inn;
        private final Map<Integer, YearGroup> yearGroups = new LinkedHashMap<>();
        private double totalValue;

        public DataGroup(String name) {
            this.name = name;
            this.inn = "";
            this.totalValue = 0;
        }

        public void addRow(DataRow row, double value) {
            Integer year = row.getYear();
            if (year == null) {
                year = 0;
            }

            // Сохраняем ИНН из строки
            if (row.getInn() != null && !row.getInn().isEmpty()) {
                this.inn = row.getInn();
            }

            YearGroup yearGroup = yearGroups.computeIfAbsent(year, YearGroup::new);
            yearGroup.addRow(row, value);
            totalValue += value;
        }

        public String getName() { return name; }
        public String getInn() { return inn; }
        public String getDisplayName() {
            if (inn != null && !inn.isEmpty()) {
                return name + " " + inn;
            }
            return name;
        }
        public Map<Integer, YearGroup> getYearGroups() { return yearGroups; }
        public double getTotalValue() { return totalValue; }
        public int getTotalRowCount() {
            return yearGroups.values().stream().mapToInt(YearGroup::getRowCount).sum();
        }

        public void sortYearGroups() {
            // Сортируем года по возрастанию
            List<Integer> sortedYears = new ArrayList<>(yearGroups.keySet());
            sortedYears.sort(Integer::compareTo);
            // Пересоздаем Map с сортировкой
            Map<Integer, YearGroup> sorted = new LinkedHashMap<>();
            for (Integer year : sortedYears) {
                sorted.put(year, yearGroups.get(year));
                if (year != 0) {
                    yearGroups.get(year).sortRowsByDate();
                }
            }
            yearGroups.clear();
            yearGroups.putAll(sorted);
        }
    }

    public static class ExcelData {
        public List<DataRow> rows;
        public Object[] headers;
        public int totalColumns;
        public List<String> sheetNames;
        public String selectedSheetName;
        public int dateColumnIndex = -1;

        public ExcelData() {
            this.rows = new ArrayList<>();
            this.sheetNames = new ArrayList<>();
        }
    }

    // ==================== МЕТОД 1: ВЫБОР ФАЙЛА И ЛИСТА ====================

    public static ExcelData selectFileAndSheet(String filePath) throws IOException {
        ExcelData data = new ExcelData();

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("Файл не существует: " + filePath);
        }

        // Явное указание типа файла
        Workbook workbook = null;
        try (FileInputStream fis = new FileInputStream(file)) {
            if (filePath.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (filePath.endsWith(".xls")) {
                workbook = new HSSFWorkbook(fis);
            } else {
                throw new IOException("Неподдерживаемый формат файла: " + filePath);
            }

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                data.sheetNames.add(workbook.getSheetName(i));
            }
            workbook.close();
        } catch (Exception e) {
            throw new IOException("Ошибка при чтении файла: " + e.getMessage(), e);
        }

        return data;
    }

    // ==================== МЕТОД 2: СОХРАНЕНИЕ ОТЧЕТА ====================

    public static void saveReportWithAllColumns(String outputFilePath,
                                                ExcelData data,
                                                String sheet1Name,
                                                String sheet2Name,
                                                int titleColumnIndex,
                                                int debitColumnIndex,
                                                int creditColumnIndex,
                                                int dateColumnIndex,
                                                int innColumnIndex,
                                                int operationColumnIndex) throws IOException {

        // Группируем данные для дебета и кредита
        Map<String, DataGroup> groupsDebit = groupByColumnWithAllData(data.rows, true, debitColumnIndex, innColumnIndex);
        Map<String, DataGroup> groupsCredit = groupByColumnWithAllData(data.rows, false, creditColumnIndex, innColumnIndex);

        // Сортируем группы и строки по датам
        for (DataGroup group : groupsDebit.values()) {
            group.sortYearGroups();
        }
        for (DataGroup group : groupsCredit.values()) {
            group.sortYearGroups();
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            // Создаем лист для дебета (только ненулевые значения)
            if (!groupsDebit.isEmpty()) {
                createFullReportSheet(workbook,
                        sheet1Name != null ? sheet1Name : "Отчет по дебету",
                        groupsDebit,
                        debitColumnIndex,
                        dateColumnIndex,
                        operationColumnIndex,
                        true);
            }

            // Создаем лист для кредита (только ненулевые значения)
            if (!groupsCredit.isEmpty()) {
                createFullReportSheet(workbook,
                        sheet2Name != null ? sheet2Name : "Отчет по кредиту",
                        groupsCredit,
                        creditColumnIndex,
                        dateColumnIndex,
                        operationColumnIndex,
                        false);
            }

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }
        }
    }

    private static Map<String, DataGroup> groupByColumnWithAllData(List<DataRow> allRows,
                                                                   boolean useDebit,
                                                                   int valueColumnIndex,
                                                                   int innColumnIndex) {
        Map<String, DataGroup> groups = new LinkedHashMap<>();

        for (DataRow row : allRows) {
            double value = useDebit ? row.getValue1() : row.getValue2();

            // Пропускаем нулевые значения
            if (value == 0) {
                continue;
            }

            String groupKey = row.getGroupName();

            DataGroup group = groups.get(groupKey);
            if (group == null) {
                group = new DataGroup(groupKey);
                groups.put(groupKey, group);
            }
            group.addRow(row, value);
        }

        return groups;
    }

    private static void createFullReportSheet(XSSFWorkbook workbook,
                                              String sheetName,
                                              Map<String, DataGroup> groups,
                                              int sumColumnIndex,
                                              int dateColumnIndex,
                                              int operationColumnIndex,
                                              boolean isDebit) {

        Sheet sheet = workbook.createSheet(sheetName);

        // Создаем стили
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle mainGroupStyle = createMainGroupHeaderStyle(workbook);
        CellStyle yearGroupStyle = createYearGroupHeaderStyle(workbook);
        CellStyle totalStyle = createTotalStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);
        CellStyle dateStyle = createDateStyle(workbook);

        int currentRow = 0;

        // Создаем заголовки (только Дата, Операция, Сумма)
        String[] headers = {"Дата", "Операция", "Сумма"};
        Row headerRow = sheet.createRow(currentRow++);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
            sheet.setColumnWidth(i, getColumnWidthForReport(i));
        }

        List<DataGroup> sortedGroups = new ArrayList<>(groups.values());
        sortedGroups.sort(Comparator.comparing(DataGroup::getName));

        for (DataGroup group : sortedGroups) {

            // Заголовок группы с ИНН
            Row mainGroupRow = sheet.createRow(currentRow++);
            Cell nameCell = mainGroupRow.createCell(0);
            String displayName = group.getDisplayName();
            nameCell.setCellValue(displayName + " (всего: " + group.getTotalRowCount() + " шт., сумма: " +
                    String.format("%,.2f", group.getTotalValue()) + ")");
            nameCell.setCellStyle(mainGroupStyle);
            sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 2));

            int mainGroupStartRow = currentRow - 1;

            // Сортируем года
            List<Integer> sortedYears = new ArrayList<>(group.getYearGroups().keySet());
            sortedYears.sort(Integer::compareTo);

            for (Integer yearInt : sortedYears) {
                YearGroup yearGroup = group.getYearGroups().get(yearInt);
                if (yearGroup == null) continue;

                // Заголовок года
                Row yearGroupRow = sheet.createRow(currentRow++);
                String yearLabel = yearGroup.getYear() == 0 ? "Без даты" : String.valueOf(yearGroup.getYear());
                Cell yearNameCell = yearGroupRow.createCell(0);
                yearNameCell.setCellValue("  " + yearLabel + " (" + yearGroup.getRowCount() + " шт., сумма: " +
                        String.format("%,.2f", yearGroup.getTotalValue()) + ")");
                yearNameCell.setCellStyle(yearGroupStyle);
                sheet.addMergedRegion(new CellRangeAddress(currentRow - 1, currentRow - 1, 0, 2));

                int yearGroupStartRow = currentRow - 1;
                int firstDetailRow = currentRow;

                // Сортируем строки по дате
                List<DataRow> sortedRows = new ArrayList<>(yearGroup.getRows());
                sortedRows.sort(Comparator.comparing(DataRow::getFullDate, Comparator.nullsLast(Comparator.naturalOrder())));

                // Детальные строки (только дата, операция, сумма)
                for (DataRow dataRow : sortedRows) {
                    Row detailRow = sheet.createRow(currentRow++);
                    Object[] rowData = dataRow.getAllData();

                    // Столбец 0: Дата (полная дата, а не только год)
                    String dateValue = "";
                    if (dateColumnIndex >= 0 && dateColumnIndex < rowData.length && rowData[dateColumnIndex] != null) {
                        dateValue = rowData[dateColumnIndex].toString();
                    } else if (dataRow.getFullDate() != null) {
                        SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");
                        dateValue = sdf.format(dataRow.getFullDate());
                    }
                    setCellValueWithStyle(detailRow.createCell(0), dateValue, dateStyle);

                    // Столбец 1: Операция (полная, без обрезания)
                    String operation = dataRow.getOperationName();
                    if (operation == null || operation.isEmpty()) {
                        if (operationColumnIndex >= 0 && operationColumnIndex < rowData.length && rowData[operationColumnIndex] != null) {
                            operation = rowData[operationColumnIndex].toString();
                        } else {
                            operation = "";
                        }
                    }
                    // Не обрезаем операцию, сохраняем полностью
                    setCellValueWithStyle(detailRow.createCell(1), operation, dataStyle);

                    // Столбец 2: Сумма
                    double sumValue = extractValueFromRow(rowData, sumColumnIndex);
                    setCellValueWithStyle(detailRow.createCell(2), sumValue, dataStyle);
                }

                int lastDetailRow = currentRow - 1;

                // Группировка строк
                if (lastDetailRow >= firstDetailRow) {
                    sheet.groupRow(firstDetailRow, lastDetailRow);
                    sheet.setRowGroupCollapsed(firstDetailRow, true);
                }

                if (lastDetailRow >= yearGroupStartRow) {
                    sheet.groupRow(yearGroupStartRow + 1, lastDetailRow);
                }
            }

            int mainGroupEndRow = currentRow - 1;
            if (mainGroupEndRow > mainGroupStartRow) {
                sheet.groupRow(mainGroupStartRow + 1, mainGroupEndRow);
                sheet.setRowGroupCollapsed(mainGroupStartRow, false);
            }

            currentRow++;
        }

        sheet.setRowSumsBelow(false);
    }

    // ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================

    private static double extractValueFromRow(Object[] rowData, int columnIndex) {
        if (columnIndex < 0 || columnIndex >= rowData.length || rowData[columnIndex] == null) {
            return 0;
        }
        try {
            if (rowData[columnIndex] instanceof Number) {
                return ((Number) rowData[columnIndex]).doubleValue();
            }
            String str = rowData[columnIndex].toString().replace(",", ".").replace(" ", "").replace(" ", "");
            if (str.isEmpty()) return 0;
            return Double.parseDouble(str);
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private static void setCellValueWithStyle(Cell cell, Object value, CellStyle style) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue(value.toString());
        }
        cell.setCellStyle(style);
    }

    private static int getColumnWidthForReport(int columnIndex) {
        switch (columnIndex) {
            case 0: return 4500;  // Дата
            case 1: return 25000; // Операция (увеличенная ширина)
            case 2: return 4500;  // Сумма
            default: return 5000;
        }
    }

    // ==================== СТИЛИ ====================

    private static CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        return style;
    }

    private static CellStyle createMainGroupHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 13);
        font.setColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        return style;
    }

    private static CellStyle createYearGroupHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.DARK_GREEN.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.LEFT);
        return style;
    }

    private static CellStyle createTotalStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createDataStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true); // Включаем перенос текста для операций
        return style;
    }

    private static CellStyle createDateStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    // ==================== МЕТОДЫ ДЛЯ ОБРАБОТКИ ДАННЫХ ====================

    public static void processDataForReport(ExcelData data,
                                            int titleColumnIndex,
                                            int debitColumnIndex,
                                            int creditColumnIndex,
                                            int dateColumnIndex,
                                            int innColumnIndex,
                                            int operationColumnIndex) {
        // Обогащаем данные: извлекаем ИНН и название операции в отдельные поля
        for (DataRow row : data.rows) {
            Object[] rowData = row.getAllData();

            // Извлекаем ИНН
            if (innColumnIndex >= 0 && innColumnIndex < rowData.length && rowData[innColumnIndex] != null) {
                String innValue = rowData[innColumnIndex].toString();
                row.setInn(extractInnFromData(innValue));
            } else {
                row.setInn("");
            }

            // Извлекаем название операции (полностью, без обрезания)
            if (operationColumnIndex >= 0 && operationColumnIndex < rowData.length && rowData[operationColumnIndex] != null) {
                row.setOperationName(rowData[operationColumnIndex].toString());
            } else {
                row.setOperationName("");
            }
        }
    }

    public static void cleanData(ExcelData data) {
        cleanData(data, true, true, true);
    }

    public static void cleanData(ExcelData data,
                                 boolean removeDuplicates,
                                 boolean removeEmptyGroups,
                                 boolean trimStrings) {

        List<DataRow> cleanedRows = new ArrayList<>();
        Set<String> seenRows = new HashSet<>();

        for (DataRow row : data.rows) {
            if (removeEmptyGroups && (row.getGroupName() == null || row.getGroupName().trim().isEmpty())) {
                continue;
            }

            if (trimStrings) {
                String trimmedName = row.getGroupName().trim();
                DataRow newRow = new DataRow(trimmedName, row.getAllData(), row.getValue1(), row.getValue2());
                newRow.setYear(row.getYear());
                newRow.setInn(row.getInn());
                newRow.setOperationName(row.getOperationName());
                if (row.getFullDate() != null) {
                    newRow.setDate(row.getFullDate());
                }
                row = newRow;
            }

            if (removeDuplicates) {
                String rowKey = row.getGroupName() + Arrays.toString(row.getAllData());
                if (!seenRows.contains(rowKey)) {
                    seenRows.add(rowKey);
                    cleanedRows.add(row);
                }
            } else {
                cleanedRows.add(row);
            }
        }

        data.rows = cleanedRows;
    }

    private static String extractInnFromData(String text) {
        if (text == null || text.isEmpty()) return "";

        // Ищем ИНН (10 или 12 цифр)
        Pattern pattern = Pattern.compile("\\b\\d{10}\\b|\\b\\d{12}\\b");
        java.util.regex.Matcher matcher = pattern.matcher(text);

        if (matcher.find()) {
            return matcher.group();
        }

        return text; // Если не нашли ИНН, возвращаем исходный текст
    }
}