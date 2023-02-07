import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;

public class ExcelFile {
    private static final List<Integer> criteriaIndex = new ArrayList<>();
    private static Integer gapIndex = -1;
    private static Integer sumIndex = -1;
    private static Integer maxIndex = -1;
    private static Integer minIndex = -1;
    private static Integer concatIndex = -1;

    public void parseExcelFile() {
        String selectedPath = "E:\\1Java\\Test\\test3.xlsx";
        String savePath = "C:\\Users\\Kirill\\Desktop\\123.xlsx";
        List<List<String>> data = getDataFromFile(selectedPath);
        List<String> firstLine = data.remove(0);
        setFirstLineIndexes(firstLine);
        System.out.println(data);
    }

    public List<List<String>> getDataFromFile(String path) {
        List<List<String>> data = new ArrayList<>();
        try {
            FileInputStream file = new FileInputStream(path);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet1 = workbook.getSheetAt(0);

            for (Row row: sheet1) {
                List<String> rowList = new ArrayList<>();
                boolean isEmptyRow = checkEmptyRow(row);
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

        Integer emptyColumnNumber = findEmptyColumnNumber(data);
        if (emptyColumnNumber == null) return data;

        return deleteEmptyColumns(data, emptyColumnNumber);
    }

    public static boolean checkEmptyRow(Row row){
        boolean isEmpty = true;
        DataFormatter dataFormatter = new DataFormatter();

        for(Cell cell: row) {
            if(dataFormatter.formatCellValue(cell).trim().length() > 0) {
                isEmpty = false;
                break;
            }
        }
        return isEmpty;
    }

    public static Integer findEmptyColumnNumber(List<List<String>> data){
        Integer emptyColumnNumber = null;

        for (int i = 0; i < data.get(0).size(); i++) {
            for (List<String> datum : data) {
                String cellValue = datum.get(i);
                if (cellValue.length() == 0) {
                    emptyColumnNumber = i;
                } else {
                    emptyColumnNumber = null;
                    break;
                }
            }
            if (emptyColumnNumber != null) break;
        }
        return emptyColumnNumber;
    }

    public static List<List<String>> deleteEmptyColumns(List<List<String>> data, int emptyColumnNumber){
        for (List<String> row : data) {
            while (row.size() > emptyColumnNumber) {
                row.remove(emptyColumnNumber);
            }
        }
        return data;
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


