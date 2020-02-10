# WorkbookEventReader
See in other languages: English | [中文](README_CN.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![Build Status](https://api.travis-ci.org/kumasuke120/workbook-event-reader.svg?branch=java8-release)](https://travis-ci.org/kumasuke120/workbook-event-reader) [![codecov](https://codecov.io/gh/kumasuke120/workbook-event-reader/branch/java8-release/graph/badge.svg)](https://codecov.io/gh/kumasuke120/workbook-event-reader)

An event-based workbook reader which re-encapsulates [Apache POI](https://poi.apache.org/) Event APIs for processing 
Excel workbooks, making it easy for reading SpreadsheetML(.xlsx) and legacy Excel documents(.xls) with a same unified interface.
All values read from an Excel document could be converted to sensible corresponding Java types.

## Branches
There are two kinds of branches in this repository. Each kind of branches have their own prefixes in their branch names:
- `java11-*`: written in Java 11, often gets newer features
- `java8-*`: written in Java 8, with more stable features (_recommended for production_)

## Requirements
Requirements for different kinds of branches may differ:

| Branch             	| Bundled 	| `java11-*`       	| `java8-*`        	|
|--------------------	|---------	|------------------	|------------------	|
| Java Ver.          	| None     	| 11 onwards    	| 8 onwards     	|
| Apache POI Ver.    	| 4.0.0   	| 3.17 onwards  	| 3.17 onwards  	|
| Apache Xerces Ver. 	| 2.12.0  	| 2.0.0 onwards 	| 2.0.0 onwards 	|

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

The Maven dependency for `java11-*` branches is:
```xml
<dependency>
    <groupId>app.kumasuke.excel</groupId>
    <artifactId>workbook-event-reader-java11</artifactId>
    <version>1.0.0</version>
</dependency>
```
The Maven dependency for `java8-*` branches is:
```xml
<dependency>
    <groupId>app.kumasuke.excel</groupId>
    <artifactId>workbook-event-reader</artifactId>
    <version>1.0.0</version>
</dependency>
```

## Sample
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