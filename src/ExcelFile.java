import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;

public class ExcelFile {
    private static final Helper helper = new Helper();
    private static final List<Integer> groupIndexes = new ArrayList<>();
    private static final List<Integer> gapIndexes = new ArrayList<>();
    private static final List<Integer> sumIndexes = new ArrayList<>();
    private static final List<Integer> maxIndexes = new ArrayList<>();
    private static final List<Integer> minIndexes = new ArrayList<>();
    private static final List<Integer> concatIndexes = new ArrayList<>();

    public void parseExcelFile(String selectedPath, String savePath) {
        // получаем данные из файла
        List<List<String>> data = getDataFromFile(selectedPath);
        // получаем первую строку и находим номера индексов для критериев и параметров группировки
        List<String> firstLine = data.remove(0);
        setFirstLineIndexes(firstLine);
        // получаем индексы строк с одинаковыми критериями группировки, например: [100, 43, 43, 43, 43, 1]=[0, 1, 7, 9] и т.д.
        Map<List<String>, List<Integer>> map = getLines(data);
        // группировка на основе данных, полученных из файла
        List<List<String>> newData = groupData(map, data);
        // создаем новый файл по указанному пути
        createNewFile(newData, savePath);
    }


    public static List<List<String>> groupData(Map<List<String>, List<Integer>> map, List<List<String>> data) {
        List<List<String>> newData = new ArrayList<>();

        for (List<Integer> lineIndexes : map.values()) {
            // если только одна строка с одинаковыми критерями группировки
            if (lineIndexes.size() == 1) {
                Integer index = lineIndexes.get(0);
                List<String> currentLine = data.get(index);
                List<String> line = getSingleLine(currentLine);
                newData.add(line);
                continue;
            }

            // если две и более строки с одинаковыми критерями группировки
            if (lineIndexes.size() > 1) {
                List<List<String>> sortedDataByLineIndexes = helper.groupDataByLineIndexes(data, lineIndexes);
                List<String> combinedLines = getCombinedLines(sortedDataByLineIndexes);
                newData.add(combinedLines);
            }
        }
        return newData;
    }

    public static List<String> getCombinedLines(List<List<String>> data) {
        List<List<String>> combinedLines = new ArrayList<>();

        for (int i = 0; i < data.size(); i++) {
            int shift = 0;
            List<String> newLine = new ArrayList<>();
            // получение текущей строки, если строк более 3, то берется  сгруппированная предыдущей итерацией строка
            List<String> currentLine = combinedLines.size() == 0 ? data.get(i) : combinedLines.get(combinedLines.size() - 1);
            List<String> nextLine = data.size() - 1 > i ? data.get(i + 1) : null;

            for (int j = 0; j < data.get(i).size(); j++) {
                boolean currentLineHasElement = helper.hasElement(currentLine, j);
                boolean nextLineHasElement = helper.hasElement(nextLine, j + shift);
                String currentLineElement = currentLineHasElement ? currentLine.get(j) : null;
                String nextLineElement = nextLine != null && nextLineHasElement ? nextLine.get(j + shift) : null;

                if (nextLineElement == null || currentLineElement == null) break;

                /* если 3 и более строки с одинаковыми критериями группировки, то учитывается свдиг,
                относительно колонки которая не участует в группировке */

                if (gapIndexes.contains(j + shift) && i >= 1) {
                    /* получаем сдвиг на основе того сколько колонок, которые не участвуют в группировке
                    идет друг за другом */
                    shift += getShift(j + shift);
                    nextLineElement = nextLine.get(j + shift);
                }

                // проверка на пустую ячейку
                if (!gapIndexes.contains(j + shift) && helper.isEmptyCell(currentLineElement, nextLineElement)) {
                    String newElement = helper.getCellContent(currentLineElement, nextLineElement);
                    if (newElement.equals("")) {
                        newLine.add(newElement);
                        continue;
                    }
                    newLine.add(helper.convertFormat(helper.getParsedNum(newElement)));
                    continue;
                }

                if (groupIndexes.contains(j + shift)) {
                    // проверка на то, является ли элемент числом
                    if (helper.iSDouble(currentLineElement)) {
                        // преобразование строки в число с обрезанием десятичной части если число целое
                        newLine.add(helper.convertFormat(helper.getParsedNum(currentLineElement)));
                    } else {
                        newLine.add(currentLineElement);
                    }
                }

                if (sumIndexes.contains(j + shift)) {
                    newLine.add(helper.convertFormat(helper.getParsedNum(currentLineElement) + helper.getParsedNum(nextLineElement)));
                }

                if (maxIndexes.contains(j + shift)) {
                    newLine.add(helper.convertFormat(Math.max(helper.getParsedNum(currentLineElement), helper.getParsedNum(nextLineElement))));
                }

                if (minIndexes.contains(j + shift)) {
                    newLine.add(helper.convertFormat(Math.min(helper.getParsedNum(currentLineElement), helper.getParsedNum(nextLineElement))));
                }

                if (concatIndexes.contains(j + shift)) {
                    if (helper.iSDouble(currentLineElement) && helper.iSDouble(nextLineElement)) {
                        newLine.add(helper.convertFormat(helper.getParsedNum(currentLineElement)) + helper.convertFormat(helper.getParsedNum(nextLineElement)));
                        continue;
                    }

                    if (helper.iSDouble(currentLineElement) && !helper.iSDouble(nextLineElement)) {
                        newLine.add(helper.convertFormat(helper.getParsedNum(currentLineElement)) + nextLineElement);
                        continue;
                    }

                    if (!helper.iSDouble(currentLineElement) && helper.iSDouble(nextLineElement)) {
                        newLine.add(currentLineElement + helper.convertFormat(helper.getParsedNum(nextLineElement)));
                        continue;
                    }

                    if (!helper.iSDouble(currentLineElement) && !helper.iSDouble(nextLineElement)) {
                        newLine.add(currentLineElement + nextLineElement);
                    }
                }
            }
            if (newLine.size() == 0) break;
            if (combinedLines.size() > 0) combinedLines.remove(0);
            combinedLines.add(newLine);
        }
        return combinedLines.get(0);
    }

    public static List<String> getSingleLine(List<String> currentLine) {
        List<String> newLine = new ArrayList<>();

        for (int i = 0; i < currentLine.size(); i++) {
            String element = currentLine.get(i);
            if (gapIndexes.contains(i)) continue;

            // проверка на пустую ячейку
            if (element.equals("")) {
                newLine.add(element);
                continue;
            }

            if (groupIndexes.contains(i) || concatIndexes.contains(i)) {
                if (helper.iSDouble(element)) {
                    newLine.add(helper.convertFormat(helper.getParsedNum(element)));
                } else {
                    newLine.add(element);
                }
            } else {
                newLine.add(helper.convertFormat(helper.getParsedNum(element)));
            }
        }
        return newLine;
    }

    public static List<List<String>> getDataFromFile(String path) {
        List<List<String>> data = new ArrayList<>();
        try {
            FileInputStream file = new FileInputStream(path);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet1 = workbook.getSheetAt(0);

            for (Row row : sheet1) {
                List<String> rowList = new ArrayList<>();
                // проверка на пустую строку
                boolean isEmptyRow = helper.checkEmptyRow(row);
                if (isEmptyRow) break;

                for (Cell cell : row) {
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

        // проверка на пустой столбец
        Integer emptyColumnNumber = helper.findEmptyColumnNumber(data);
        if (emptyColumnNumber == null) return data;

        // если есть пустой столбец, то обрезаем его и лишние столбцы
        return helper.deleteEmptyColumns(data, emptyColumnNumber);
    }

    public static void createNewFile(List<List<String>> data, String path) {
        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet();

            for (int i = 0; i < data.size(); i++) {
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

    public static void setFirstLineIndexes(List<String> firstLine) {
        for (int i = 0; i < firstLine.size(); i++) {
            String element = firstLine.get(i);
            if ("".equals(element)) groupIndexes.add(i);
            if ("-".equals(element)) gapIndexes.add(i);
            if ("SUM".equals(element)) sumIndexes.add(i);
            if ("MAX".equals(element)) maxIndexes.add(i);
            if ("MIN".equals(element)) minIndexes.add(i);
            if ("CONCAT".equals(element)) concatIndexes.add(i);
        }
    }

    public static Map<List<String>, List<Integer>> getLines(List<List<String>> data) {
        Map<List<String>, List<Integer>> map = new LinkedHashMap<>();

        for (int i = 0; i < data.size(); i++) {
            List<String> line = data.get(i);
            List<String> keys = new ArrayList<>();
            List<Integer> values = new ArrayList<>();

            for (Integer index : groupIndexes) {
                keys.add(line.get(index));
            }

            if (map.containsKey(keys)) values.addAll(map.get(keys));

            values.add(i);
            map.put(keys, values);
        }
        return map;
    }

    public static int getShift(Integer index) {
        int shift = 1;
        int gapIndex = gapIndexes.indexOf(index);

        for (int i = gapIndex; i < gapIndexes.size(); i++) {
            Integer gapElement = gapIndexes.get(i);
            Integer nextGapElement = gapIndexes.size() - 1 > i ? gapIndexes.get(i + 1) : null;

            if (nextGapElement == null) break;

            if (nextGapElement - gapElement > 1) {
                break;
            }

            if (nextGapElement - gapElement == 1) {
                shift++;
            }
        }
        return shift;
    }
}
