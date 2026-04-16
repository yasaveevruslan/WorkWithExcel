package RoboSimJava;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.Formula;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class FunctionExcel {

    public static List<String> readSheet(String name) {
        List<String> sheetNames = new ArrayList<>();
        File file = new File(name);

        if (!file.exists()) {
            return sheetNames;
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return sheetNames;
    }

    public static void read(String nameList, String name, Map<String, Object[]> dates) throws IOException, InvalidFormatException {
        try (InputStream inp = new FileInputStream(name)) {
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = nameList == null || nameList.isEmpty() ? wb.getSheetAt(0) : wb.getSheet(nameList);

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            int i = 1;
            for (Row row : sheet) {
                Object[] ob = new Object[row.getLastCellNum()];
                for (int j = 0; j < row.getLastCellNum(); j++){
                    Cell cell = row.getCell(j);
                    ob[j] = cell == null ? "" : getCellValue(cell, formatter, evaluator);
                }
                dates.put("" + i, ob);
                i++;
            }
        }
    }
    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("dd.MM.yyyy");

    private static Object getCellValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();

            case NUMERIC: {
                if (DateUtil.isCellDateFormatted(cell)) {
                    // ФОРМАТИРУЕМ ДАТУ В СТРОКУ
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
                            // Проверяем, не дата ли это в формуле
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

    public static void saveDateInExcel(String nameList, String name, Map<String, Object[]> dates) {
        XSSFWorkbook workbook;
        File file = new File(name);

        if (file.exists()) {
            try (FileInputStream fis = new FileInputStream(file)) {
                workbook = new XSSFWorkbook(fis);
            } catch (IOException e) {
                workbook = new XSSFWorkbook();
            }
        } else {
            workbook = new XSSFWorkbook();
        }

        if (nameList.isEmpty()) {
            nameList = "Лист";
        }
        String uniqueSheetName = getUniqueSheetName(workbook, nameList);
        XSSFSheet newSheet = workbook.createSheet(uniqueSheetName);
        writeDataToSheet(newSheet, dates);



        try (FileOutputStream out = new FileOutputStream(name)) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String) {
            String str = (String) value;
            if (str.startsWith("=")) {
                cell.setCellFormula(str.substring(1));
            } else {
                cell.setCellValue(str);
            }
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

    private static String getUniqueSheetName(XSSFWorkbook workbook, String baseName) {
        if (workbook.getSheet(baseName) == null) {
            return baseName;
        }

        int counter = 1;
        String newName;
        do {
            newName = baseName + " (" + counter + ")";
            counter++;
        } while (workbook.getSheet(newName) != null);

        return newName;
    }

    private static void writeDataToSheet(XSSFSheet sheet, Map<String, Object[]> data) {
        int rowNum = 0;
        for (String key : data.keySet()) {
            Row row = sheet.createRow(rowNum++);
            Object[] objArr = data.get(key);
            int cellNum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellNum++);
                setCellValue(cell, obj);
            }
        }
    }
}
