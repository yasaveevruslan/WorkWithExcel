package RoboSimJava;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelGrouper {

    // ==================== ВСПОМОГАТЕЛЬНЫЕ КЛАССЫ ====================

    public static class DataRow {
        private String groupName;
        private final Object[] allData;
        private final double value1;
        private final double value2;
        private Integer year;

        public DataRow(String groupName, Object[] allData, double value1, double value2) {
            this.groupName = groupName;
            this.allData = allData;
            this.value1 = value1;
            this.value2 = value2;
        }

        public String getGroupName() { return groupName; }
        public Object[] getAllData() { return allData; }
        public double getValue1() { return value1; }
        public double getValue2() { return value2; }
        public Integer getYear() { return year; }

        public void setDate(Date date) {
            if (date != null) {
                Calendar cal = Calendar.getInstance();
                cal.setTime(date);
                this.year = cal.get(Calendar.YEAR);
            }
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
    }

    public static class DataGroup {
        private final String name;
        private final Map<Integer, YearGroup> yearGroups = new LinkedHashMap<>();
        private double totalValue;

        public DataGroup(String name) {
            this.name = name;
            this.totalValue = 0;
        }

        public void addRow(DataRow row, double value) {
            Integer year = row.getYear();
            if (year == null) {
                year = 0;
            }

            YearGroup yearGroup = yearGroups.computeIfAbsent(year, YearGroup::new);
            yearGroup.addRow(row, value);
            totalValue += value;
        }

        public String getName() { return name; }
        public Map<Integer, YearGroup> getYearGroups() { return yearGroups; }
        public double getTotalValue() { return totalValue; }
        public int getTotalRowCount() {
            return yearGroups.values().stream().mapToInt(YearGroup::getRowCount).sum();
        }
        public int getYearGroupCount() { return yearGroups.size(); }
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

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                data.sheetNames.add(workbook.getSheetName(i));
            }
        }
        return data;
    }

    public static void selectSheetByName(ExcelData data, String sheetName) {
        if (data.sheetNames.contains(sheetName)) {
            data.selectedSheetName = sheetName;
        }
    }

    // ==================== МЕТОД 2: ЧТЕНИЕ ВЫБРАННОГО ЛИСТА ====================

    public static void readSelectedSheet(String filePath,
                                         ExcelData data,
                                         int groupNameColumnIndex,
                                         int valueColumn1Index,
                                         int valueColumn2Index,
                                         int dateColumnIndex) throws IOException {

        data.dateColumnIndex = dateColumnIndex;

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet;
            if (data.selectedSheetName != null) {
                sheet = workbook.getSheet(data.selectedSheetName);
            } else {
                sheet = workbook.getSheetAt(0);
                data.selectedSheetName = sheet.getSheetName();
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            Row firstRow = sheet.getRow(0);
            data.totalColumns = firstRow != null ? firstRow.getLastCellNum() : 10;

            boolean hasHeader = isHeaderRow(sheet.getRow(0));
            int startRow = hasHeader ? 1 : 0;

            if (hasHeader) {
                Row headerRow = sheet.getRow(0);
                data.headers = new Object[data.totalColumns];
                for (int j = 0; j < data.totalColumns; j++) {
                    Cell cell = headerRow.getCell(j);
                    data.headers[j] = getCellValue(cell, formatter, evaluator);
                }
            }

            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell nameCell = row.getCell(groupNameColumnIndex);
                String groupName = getCellValueAsString(nameCell, formatter, evaluator);
                if (groupName.trim().isEmpty()) continue;

                Object[] rowData = new Object[data.totalColumns];
                for (int j = 0; j < data.totalColumns; j++) {
                    Cell cell = row.getCell(j);
                    rowData[j] = getCellValue(cell, formatter, evaluator);
                }

                double value1 = getNumericValue(row.getCell(valueColumn1Index), formatter, evaluator);
                double value2 = getNumericValue(row.getCell(valueColumn2Index), formatter, evaluator);

                DataRow dataRow = new DataRow(groupName, rowData, value1, value2);

                if (dateColumnIndex >= 0) {
                    Cell dateCell = row.getCell(dateColumnIndex);
                    Date date = getDateValue(dateCell, formatter, evaluator);
                    dataRow.setDate(date);
                }

                data.rows.add(dataRow);
            }
        }
    }

    public static void readSelectedSheetWithProgress(String filePath,
                                                     ExcelData data,
                                                     int groupNameColumnIndex,
                                                     int valueColumn1Index,
                                                     int valueColumn2Index,
                                                     int dateColumnIndex,
                                                     ProgressDialog progressDialog) throws IOException, InterruptedException {

        data.dateColumnIndex = dateColumnIndex;
        data.rows.clear();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet;
            if (data.selectedSheetName != null) {
                sheet = workbook.getSheet(data.selectedSheetName);
            } else {
                sheet = workbook.getSheetAt(0);
                data.selectedSheetName = sheet.getSheetName();
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            Row firstRow = sheet.getRow(0);
            data.totalColumns = firstRow != null ? firstRow.getLastCellNum() : 10;

            boolean hasHeader = isHeaderRow(sheet.getRow(0));
            int startRow = hasHeader ? 1 : 0;
            int totalRows = sheet.getLastRowNum() - startRow + 1;

            if (hasHeader) {
                Row headerRow = sheet.getRow(0);
                data.headers = new Object[data.totalColumns];
                for (int j = 0; j < data.totalColumns; j++) {
                    Cell cell = headerRow.getCell(j);
                    data.headers[j] = getCellValue(cell, formatter, evaluator);
                }
            }

            int processedRows = 0;
            int lastProgress = -1;

            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
                if (progressDialog != null && progressDialog.isCancelled()) {
                    throw new InterruptedException("Операция отменена пользователем");
                }

                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell nameCell = row.getCell(groupNameColumnIndex);
                String groupName = getCellValueAsString(nameCell, formatter, evaluator);
                if (groupName.trim().isEmpty()) continue;

                Object[] rowData = new Object[data.totalColumns];
                for (int j = 0; j < data.totalColumns; j++) {
                    Cell cell = row.getCell(j);
                    rowData[j] = getCellValue(cell, formatter, evaluator);
                }

                double value1 = getNumericValue(row.getCell(valueColumn1Index), formatter, evaluator);
                double value2 = getNumericValue(row.getCell(valueColumn2Index), formatter, evaluator);

                DataRow dataRow = new DataRow(groupName, rowData, value1, value2);

                if (dateColumnIndex >= 0) {
                    Cell dateCell = row.getCell(dateColumnIndex);
                    Date date = getDateValue(dateCell, formatter, evaluator);
                    dataRow.setDate(date);
                }

                data.rows.add(dataRow);
                processedRows++;

                if (progressDialog != null && processedRows % 50 == 0) {
                    int progress = (int) ((double) processedRows / totalRows * 100);
                    if (progress != lastProgress) {
                        progressDialog.setProgress(Math.min(progress, 100));
                        progressDialog.setStatus("Обработка строки " + processedRows + " из " + totalRows);
                        lastProgress = progress;
                    }
                }
            }

            if (progressDialog != null && !progressDialog.isCancelled()) {
                progressDialog.setProgress(100);
                progressDialog.setStatus("Обработка завершена!");
            }
        }
    }

    // ==================== МЕТОД 3: ОЧИСТКА ДАННЫХ ====================

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
            if (removeEmptyGroups && (row.groupName == null || row.groupName.trim().isEmpty())) {
                continue;
            }

            if (trimStrings) {
                row.groupName = row.groupName.trim();
                for (int i = 0; i < row.allData.length; i++) {
                    if (row.allData[i] instanceof String) {
                        row.allData[i] = ((String) row.allData[i]).trim();
                    }
                }
            }

            if (removeDuplicates) {
                String rowKey = row.groupName + Arrays.toString(row.allData);
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

    // ==================== МЕТОД 4: СОХРАНЕНИЕ ФАЙЛА С ГРУППИРОВКОЙ ПО ГОДАМ ====================

    public static void saveGroupedFileWithYears(String outputFilePath,
                                                ExcelData data,
                                                String sheet1Name,
                                                String sheet2Name,
                                                int valueColumn1Index,
                                                int valueColumn2Index) throws IOException {

        Map<String, DataGroup> groups1 = groupByColumnWithYears(data.rows, true);
        Map<String, DataGroup> groups2 = groupByColumnWithYears(data.rows, false);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {

            createGroupedSheetWithYears(workbook,
                    sheet1Name != null ? sheet1Name : "Столбец 1 (все)",
                    groups1, data.headers, valueColumn1Index);

            createGroupedSheetWithYears(workbook,
                    sheet2Name != null ? sheet2Name : "Столбец 2 (без нулей)",
                    groups2, data.headers, valueColumn2Index);

            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }

        }
    }

    private static Map<String, DataGroup> groupByColumnWithYears(List<DataRow> allRows, boolean useColumn1) {
        Map<String, DataGroup> groups = new LinkedHashMap<>();

        List<DataRow> sortedRows = new ArrayList<>(allRows);
        sortedRows.sort(Comparator.comparing(DataRow::getGroupName)
                .thenComparing(DataRow::getYear, Comparator.nullsLast(Comparator.naturalOrder())));

        for (DataRow row : sortedRows) {
            double value = useColumn1 ? row.getValue1() : row.getValue2();

            // Фильтруем нули для всех листов (и для дебета, и для кредита)
            if (value == 0) {
                continue;
            }

            DataGroup group = groups.computeIfAbsent(row.getGroupName(), DataGroup::new);
            group.addRow(row, value);
        }

        return groups;
    }

    private static void createGroupedSheetWithYears(XSSFWorkbook workbook,
                                                    String sheetName,
                                                    Map<String, DataGroup> groups,
                                                    Object[] headers,
                                                    int valueColumnIndex) {

        Sheet sheet = workbook.createSheet(sheetName);

        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle mainGroupStyle = createMainGroupHeaderStyle(workbook);
        CellStyle yearGroupStyle = createYearGroupHeaderStyle(workbook);
        CellStyle totalStyle = createTotalStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);

        int currentRow = 0;

        if (headers != null) {
            Row headerRow = sheet.createRow(currentRow++);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i] != null ? headers[i].toString() : "");
                cell.setCellStyle(headerStyle);
            }
        }

        List<DataGroup> sortedGroups = new ArrayList<>(groups.values());
        sortedGroups.sort(Comparator.comparing(DataGroup::getName));

        for (DataGroup group : sortedGroups) {

            Row mainGroupRow = sheet.createRow(currentRow++);

            Cell nameCell = mainGroupRow.createCell(0);
            nameCell.setCellValue(group.getName() + " (всего: " + group.getTotalRowCount() + " шт.)");
            nameCell.setCellStyle(mainGroupStyle);

            Cell totalCell = mainGroupRow.createCell(valueColumnIndex);
            totalCell.setCellValue(group.getTotalValue());
            totalCell.setCellStyle(totalStyle);

            int mainGroupStartRow = currentRow - 1;

            List<YearGroup> sortedYearGroups = new ArrayList<>(group.getYearGroups().values());
            sortedYearGroups.sort(Comparator.comparing(YearGroup::getYear));

            for (YearGroup yearGroup : sortedYearGroups) {

                Row yearGroupRow = sheet.createRow(currentRow++);

                Cell yearNameCell = yearGroupRow.createCell(1);
                String yearLabel = yearGroup.getYear() == 0 ? "Без даты" : String.valueOf(yearGroup.getYear());
                yearNameCell.setCellValue(yearLabel + " (" + yearGroup.getRowCount() + " шт.)");
                yearNameCell.setCellStyle(yearGroupStyle);

                Cell yearTotalCell = yearGroupRow.createCell(valueColumnIndex);
                yearTotalCell.setCellValue(yearGroup.getTotalValue());
                yearTotalCell.setCellStyle(totalStyle);

                int yearGroupStartRow = currentRow - 1;
                int firstDetailRow = currentRow;

                for (DataRow dataRow : yearGroup.getRows()) {
                    Row detailRow = sheet.createRow(currentRow++);
                    Object[] rowData = dataRow.getAllData();

                    for (int i = 0; i < rowData.length; i++) {
                        Cell cell = detailRow.createCell(i);
                        setCellValue(cell, rowData[i]);
                        cell.setCellStyle(dataStyle);
                    }
                }

                int lastDetailRow = currentRow - 1;

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

        if (headers != null) {
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }
        }

        sheet.setRowSumsBelow(false);
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
        return style;
    }

    private static CellStyle createMainGroupHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.MEDIUM);
        return style;
    }

    private static CellStyle createYearGroupHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.BLUE.getIndex());
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
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
        return style;
    }

    private static CellStyle createDataStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    // ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================

    private static Date getDateValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) return null;

        try {
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            }

            String strValue = formatter.formatCellValue(cell, evaluator);
            if (strValue == null || strValue.trim().isEmpty()) return null;

            String[] patterns = {"dd.MM.yyyy", "yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy"};
            for (String pattern : patterns) {
                try {
                    SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                    return sdf.parse(strValue);
                } catch (Exception e) {
                    // Пробуем следующий формат
                }
            }
        } catch (Exception e) {
            // Игнорируем
        }
        return null;
    }

    private static boolean isHeaderRow(Row row) {
        if (row == null) return false;
        Cell cell = row.getCell(0);
        return cell != null && cell.getCellType() == CellType.STRING;
    }

    private static String getCellValueAsString(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) return "";
        return formatter.formatCellValue(cell, evaluator).trim();
    }

    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("dd.MM.yyyy");

    private static Object getCellValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();

            case NUMERIC: {
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    return DATE_FORMAT.format(date);
                }
                double num = cell.getNumericCellValue();
                return num == (long) num ? (long) num : num;
            }
            case BOOLEAN:
                return cell.getBooleanCellValue();

            case FORMULA: {
                try {
                    CellValue cv = evaluator.evaluate(cell);
                    switch (cv.getCellType()) {
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                Date date = cell.getDateCellValue();
                                return DATE_FORMAT.format(date);
                            }
                            double num = cv.getNumberValue();
                            return num == (long) num ? (long) num : num;
                        case STRING:
                            return cv.getStringValue();
                        case BOOLEAN:
                            return cv.getBooleanValue();
                        default:
                            return cell.getCellFormula();
                    }
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            }
            case BLANK:
                return "";

            default:
                return formatter.formatCellValue(cell);
        }
    }

    private static double getNumericValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) return 0;

        try {
            String str = formatter.formatCellValue(cell, evaluator);
            if (str == null || str.trim().isEmpty()) return 0;
            return Double.parseDouble(str.replace(",", ".").replace(" ", ""));
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    private static void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }

    
}