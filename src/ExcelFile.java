import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.*;
import java.util.*;

public class ExcelFile {
    private static final Helper helper = new Helper();
    private static final List<Integer> criteriaIndex = new ArrayList<>();
    private static Integer gapIndex = -1;
    private static Integer sumIndex = -1;
    private static Integer maxIndex = -1;
    private static Integer minIndex = -1;
    private static Integer concatIndex = -1;

    public void parseExcelFile(String selectedPath, String savePath ) {
        List<List<String>> data = getDataFromFile(selectedPath);
        List<String> firstLine = data.remove(0);
        setFirstLineIndexes(firstLine);
        List<List<String>> newData = sortData(data);
        createNewFile(newData, savePath);
    }

    public static List<List<String>> getDataFromFile(String path) {
        List<List<String>> data = new ArrayList<>();
        try {
            FileInputStream file = new FileInputStream(path);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet1 = workbook.getSheetAt(0);

            for (Row row: sheet1) {
                List<String> rowList = new ArrayList<>();
                boolean isEmptyRow = helper.checkEmptyRow(row);
                if (isEmptyRow) break;

                for (Cell cell: row) {
                    switch (cell.getCellType()) {
                        case BLANK, STRING -> rowList.add(cell.getStringCellValue());
                        case NUMERIC -> rowList.add(String.valueOf(cell.getNumericCellValue()));
                    }
                }
                data.add(rowList);
            }
        } catch (Exception e) {
            System.out.println("Что-то пошло не так! " + e.getMessage());
        }

        Integer emptyColumnNumber = helper.findEmptyColumnNumber(data);
        if (emptyColumnNumber == null) return data;

        return helper.deleteEmptyColumns(data, emptyColumnNumber);
    }

    public static void createNewFile(List<List<String>> data, String path){
        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet();

            for (int i =0; i < data.size(); i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < data.get(i).size(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(data.get(i).get(j));
                }
            }

            FileOutputStream outputStream = new FileOutputStream(path);
            workbook.write(outputStream);
            workbook.close();
        } catch (Exception e) {
            System.out.println("Что-то пошло не так! " + e.getMessage());
        }
    }

    public static List<List<String>> sortData(List<List<String>> data) {
        List<List<String>> sortedData = new ArrayList<>();
        String element;
        String nextRowElement;
        boolean flag = false;

        for (int i = 0; i < data.size(); i++) {
            List<String> line = new ArrayList<>();
            String lineFirstElement = data.get(i).get(0);
            String lineSecondElement = data.get(i).get(1);

            if (flag) {
                flag = false;
                continue;
            }

            for (int j = 0; j < data.get(i).size(); j++) {
                element = data.get(i).get(j);
                nextRowElement = data.size() - 1 > i ? data.get(i + 1).get(j) : element;

                if (data.size() - 1 > i &&
                        lineFirstElement.equals(data.get(i + 1).get(0)) &&
                        lineSecondElement.equals(data.get(i + 1).get(1))) {

                    if (j != gapIndex && helper.isEmptyCell(element, nextRowElement)) {
                        String newElement = helper.getCellContent(element, nextRowElement);
                        if (newElement.equals("")) {
                            line.add(newElement);
                            continue;
                        }
                        line.add(helper.convertFormat(helper.getParsedNum(newElement)));
                        continue;
                    }

                    if (criteriaIndex.contains(j)) {
                        if (helper.iSDouble(element)) {
                            line.add(helper.convertFormat(helper.getParsedNum(element)));
                        } else {
                            line.add(element);
                        }
                    }

                    if (j == sumIndex) {
                        line.add(helper.convertFormat(helper.getParsedNum(element) + helper.getParsedNum(nextRowElement)));
                    }

                    if (j == maxIndex) {
                        line.add(helper.convertFormat(Math.max(helper.getParsedNum(element), helper.getParsedNum(nextRowElement))));
                    }

                    if (j == minIndex) {
                        line.add(helper.convertFormat(Math.min(helper.getParsedNum(element), helper.getParsedNum(nextRowElement))));
                    }

                    if (j == concatIndex) {
                        if (helper.iSDouble(element)) {
                            line.add(helper.convertFormat(helper.getParsedNum(element)) + helper.convertFormat(helper.getParsedNum(nextRowElement)));
                        } else {
                            line.add(element + nextRowElement);
                        }
                    }
                    flag = true;
                } else {
                    if (j == gapIndex) continue;

                    if (element.equals("")) {
                        line.add(element);
                        continue;
                    }

                    if (criteriaIndex.contains(j) || j == concatIndex) {
                        if (helper.iSDouble(element)) {
                            line.add(helper.convertFormat(helper.getParsedNum(element)));
                        } else {
                            line.add(element);
                        }
                    } else {
                        line.add(helper.convertFormat(helper.getParsedNum(element)));
                    }
                    flag = false;
                }
            }
            sortedData.add(line);
        }
        return sortedData;
    }

    public static void setFirstLineIndexes(List<String> firstLine) {
        for (int i = 0; i < firstLine.size(); i++) {
            String element = firstLine.get(i);
            if ("".equals(element)) criteriaIndex.add(i);
            if ("-".equals(element)) gapIndex = i;
            if ("SUM".equals(element)) sumIndex = i;
            if ("MAX".equals(element)) maxIndex = i;
            if ("MIN".equals(element)) minIndex = i;
            if ("CONCAT".equals(element)) concatIndex = i;
        }
    }
}
