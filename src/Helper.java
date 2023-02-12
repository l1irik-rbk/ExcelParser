import org.apache.poi.ss.usermodel.*;

import java.text.*;
import java.util.*;

public class Helper {
    public boolean hasElement(List<String> currentLine, Integer index) {
        try {
            currentLine.get(index);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    public List<List<String>> groupDataByLineIndexes(List<List<String>> data, List<Integer> lineIndexes) {
        List<List<String>> groupedData = new ArrayList<>();

        for (Integer index : lineIndexes) {
            groupedData.add(data.get(index));
        }
        return groupedData;
    }

    public double getParsedNum(String element) {
        return Double.parseDouble(element);
    }

    public boolean isEmptyCell(String element, String nextRowElement) {
        return element.equals("") || nextRowElement.equals("");
    }

    public String getCellContent(String element, String nextRowElement) {
        if (element.equals("")) {
            return nextRowElement;
        }
        return element;
    }

    public boolean iSDouble(String str) {
        try {
            Double.parseDouble(str);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    public String convertFormat(Double num) {
        DecimalFormat format = new DecimalFormat("#.###");
        format.setDecimalSeparatorAlwaysShown(false);
        return format.format(num);
    }

    public boolean checkEmptyRow(Row row) {
        boolean isEmpty = true;
        DataFormatter dataFormatter = new DataFormatter();

        for (Cell cell : row) {
            if (dataFormatter.formatCellValue(cell).trim().length() > 0) {
                isEmpty = false;
                break;
            }
        }
        return isEmpty;
    }

    public Integer findEmptyColumnNumber(List<List<String>> data) {
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

    public List<List<String>> deleteEmptyColumns(List<List<String>> data, int emptyColumnNumber) {
        for (List<String> row : data) {
            while (row.size() > emptyColumnNumber) {
                row.remove(emptyColumnNumber);
            }
        }
        return data;
    }
}
