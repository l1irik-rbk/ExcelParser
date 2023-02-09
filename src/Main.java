public class Main {
    public static void main(String[] args) {
        // запуск основного окна
        String selectedPath = "E:\\1Java\\Test\\test4.xlsx";
        String savePath = "C:\\Users\\Kirill\\Desktop\\123.xlsx";

//        new ExcelParser();
        new ExcelFile().parseExcelFile(selectedPath, savePath);
    }
}