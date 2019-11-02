package app.kumasuke.demo;

import app.kumasuke.excel.CellValue;
import app.kumasuke.excel.WorkbookEventReader;
import app.kumasuke.excel.XSSFWorkbookEventReader;
import app.kumasuke.util.ResourceUtil;

import java.nio.file.Path;

public class DocumentPrinter {

    public static void main(String[] args) {
        System.out.println("Print 'workbook.xlsx':");
        System.out.println("--------------------------------------------------");
        readAndPrint("workbook.xlsx");

        System.out.println("--------------------------------------------------");
        System.out.println();

        System.out.println("Print 'workbook.xls':");
        System.out.println("--------------------------------------------------");
        readAndPrint("workbook.xls");
    }

    private static void readAndPrint(String fileName) {
        final Path filePath = ResourceUtil.getPathOfClasspathResource(fileName);
        try (final var reader = WorkbookEventReader.open(filePath)) {
            if (reader instanceof XSSFWorkbookEventReader) {
                ((XSSFWorkbookEventReader) reader).setUse1904Windowing(true);
            }
            reader.read(MyHandler.INSTANCE);
        }
    }

    private static class MyHandler implements WorkbookEventReader.EventHandler {
        private static final WorkbookEventReader.EventHandler INSTANCE = new MyHandler();

        @Override
        public void onStartDocument() {
            System.out.println(">> Document");
        }

        @Override
        public void onEndDocument() {
            System.out.println("<< Document");
        }

        @Override
        public void onStartSheet(int sheetIndex, String sheetName) {
            System.out.printf(">> Sheet (%d, %s)%n", sheetIndex, sheetName);
        }

        @Override
        public void onEndSheet(int sheetIndex) {
            System.out.printf("<< Sheet (%d)%n", sheetIndex);
        }

        @Override
        public void onStartRow(int sheetIndex, int rowNum) {
            System.out.printf(">> Row (%d, %d)%n", sheetIndex, rowNum);
        }

        @Override
        public void onEndRow(int sheetIndex, int rowNum) {
            System.out.printf("<< Row (%d, %d)%n", sheetIndex, rowNum);
        }

        @Override
        public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
            String cellType = cellValue.isNull() ? "null" : cellValue.originalType().getSimpleName();
            System.out.printf("   Cell (%d, %d, %d, [%s] %s)%n", sheetIndex, rowNum, columnNum,
                              cellType, cellValue.originalValue());
        }
    }

}
