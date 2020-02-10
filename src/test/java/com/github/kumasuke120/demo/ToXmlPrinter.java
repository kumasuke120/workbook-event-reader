package com.github.kumasuke120.demo;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.util.ResourceUtil;

import java.nio.file.Path;

public class ToXmlPrinter {

    public static void main(String[] args) {
        final Path filePath = ResourceUtil.getPathOfClasspathResource("workbook.xlsx");
        final XmlGenerator xmlGenerator = new XmlGenerator();
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            reader.read(xmlGenerator);
        }

        final String theXml = xmlGenerator.getXml();
        System.out.println(theXml);
    }

    private static class XmlGenerator implements WorkbookEventReader.EventHandler {
        private final StringBuilder xml;

        private int currentIndentLevel = 0;

        XmlGenerator() {
            xml = new StringBuilder();
        }

        String getXml() {
            return xml.toString();
        }

        @Override
        public void onStartDocument() {
            xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");

            newLine();
            xml.append("<document>");
            indent();
        }

        @Override
        public void onEndDocument() {
            unindent();
            newLine();
            xml.append("</document>");
        }

        @Override
        public void onStartSheet(int sheetIndex, String sheetName) {
            newLine();
            xml.append("<sheet index=\"")
                    .append(sheetIndex)
                    .append("\" name=\"")
                    .append(sheetName)
                    .append("\">");
            indent();
        }

        @Override
        public void onEndSheet(int sheetIndex) {
            unindent();
            newLine();
            xml.append("</sheet>");
        }

        @Override
        public void onStartRow(int sheetIndex, int rowNum) {
            newLine();
            xml.append("<row index=\"")
                    .append(rowNum)
                    .append("\">");
            indent();
        }

        @Override
        public void onEndRow(int sheetIndex, int rowNum) {
            unindent();
            newLine();
            xml.append("</row>");
        }

        @Override
        public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
            newLine();
            xml.append("<cell index=\"")
                    .append(columnNum)
                    .append("\" javaType=\"")
                    .append(cellValue.isNull() ? "null" : cellValue.originalType().getCanonicalName())
                    .append("\"");
            if (cellValue.isNull()) {
                xml.append("/>");
            } else {
                xml.append(">")
                        .append(cellValue.isNull() ? "" : cellValue.originalValue())
                        .append("</cell>");
            }
        }

        private void newLine() {
            xml.append(System.lineSeparator());
            xml.append(repeatFourSpaces(currentIndentLevel));
        }

        private String repeatFourSpaces(int times) {
            final StringBuilder sb = new StringBuilder();
            for (int i = 0; i < times; i++) {
                sb.append(/* four spaces */"    ");
            }
            return sb.toString();
        }

        private void indent() {
            currentIndentLevel += 1;
        }

        private void unindent() {
            currentIndentLevel -= 1;
        }
    }

}
