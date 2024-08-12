# WorkbookEventReader
See in other languages: English | [中文](README_CN.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![Build Status](https://github.com/kumasuke120/workbook-event-reader/actions/workflows/build.yml/badge.svg)](https://github.com/kumasuke120/workbook-event-reader/actions/workflows/build.yml) [![codecov](https://codecov.io/gh/kumasuke120/workbook-event-reader/branch/master/graph/badge.svg)](https://codecov.io/gh/kumasuke120/workbook-event-reader)

## Features
- Event-based Workbook reader, with a low memory footprint
- Unified interfaces for reading SpreadsheetML(.xlsx), legacy Excel documents(.xls) and CSV files
- Able to read a file from scratch for multiple times
- Easy to read the cell values, type-converting them at ease 

## Supported Documents
- *.xlsx
- *.xls
- *.csv (auto-detecting the file encoding by default)

## Requirements
Requirements for this project are as follows:

| Dependency               	 | Bundled 	  | Minimum  	 |
|----------------------------|------------|------------|
| Java Ver.             	    | None     	 | 8      	   |
| Apache POI Ver.       	    | 4.1.2   	  | 3.17  	    |
| Apache Xerces Ver.    	    | 2.12.2  	  | 2.0.0 	    |
| Apache Commons CSV Ver. 	  | 1.8     	  | 1.0   	    |

_Click [here](https://bz.apache.org/bugzilla/show_bug.cgi?id=61034) to take a look at the reason 
for the requirement of Apache POI 3.17 onwards_

## Build and Installation
You could use the following command to install:
```
$ ./mvnw clean install -DskipTests
```
or you could use your own maven to install:
```
$ mvn clean install -DskipTests
```

The Maven dependency for this project is:
```xml
<dependency>
    <groupId>com.github.kumasuke120</groupId>
    <artifactId>workbook-event-reader</artifactId>
    <version>1.5.1</version>
</dependency>
```

## Quick Start
The following code reads a workbook and converts its content to a well-formed XML document:
```java
public class ToXmlPrinter {

    public static void main(String[] args) {
        final Path filePath = Paths.get("workbook.xlsx");
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
                    .append("\">")
                    .append(cellValue.isNull() ? "" : cellValue.originalValue())
                    .append("</cell>");
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
``` 
