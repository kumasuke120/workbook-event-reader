# WorkbookEventReader
See in other languages: English | [中文](README_CN.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![Build Status](https://github.com/kumasuke120/workbook-event-reader/actions/workflows/build.yml/badge.svg)](https://github.com/kumasuke120/workbook-event-reader/actions/workflows/build.yml) [![codecov](https://codecov.io/gh/kumasuke120/workbook-event-reader/branch/master/graph/badge.svg)](https://codecov.io/gh/kumasuke120/workbook-event-reader)

## Features
- Using handy annotations to read Excel files, converting them to Java objects
- Event-based workbook reader with unified interfaces for SpreadsheetML (.xlsx), legacy Excel (.xls), and CSV files
- Efficiently processes large workbook files with a low memory footprint
- Capable of reading files multiple times from scratch, offering convenient cell value retrieval and type conversions

## Supported Documents
- *.xlsx
- *.xls
- *.csv (auto-detecting the file encoding by default)

## Requirements
Requirements for this project are as follows:

| Dependency               	 | Bundled 	  | Minimum  	                                                                   |
|----------------------------|------------|------------------------------------------------------------------------------|
| Java Ver.             	    | None     	 | 8      	                                                                     |
| Apache POI Ver.       	    | 4.1.2   	  | 3.17<sup>[[?]](https://bz.apache.org/bugzilla/show_bug.cgi?id=61034)</sup> 	 |
| Apache Xerces Ver.    	    | 2.12.2  	  | 2.0.0 	                                                                      |
| Apache Commons CSV Ver. 	  | 1.8     	  | 1.0   	                                                                      |

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
    <version>1.6.1</version>
</dependency>
```

## Quick Start
### Annotation Reading
With the help of annotations, you can use the following code to convert workbook contents to Java Objects and print its content:

```java
public class AnnotationObjectReadPrinter {

    public static void main(String[] args) {
        final Path filePath = Paths.get("workbook.xlsx");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            // extracts the records to Java Objects
            final WorkbookRecordExtractor<OrderDetail> extractor = WorkbookRecordExtractor.ofRecord(OrderDetail.class);
            reader.read(extractor);

            // prints the result
            printResult(extractor);
        }
    }

    private static void printResult(WorkbookRecordExtractor<OrderDetail> extractor) {
        printOrderDetailTitle(extractor.getColumnTitles());
        for (OrderDetail orderDetail : extractor.getResult()) {
            printOrderDetail(orderDetail);
        }
    }

    private static void printOrderDetailTitle(List<String> titles) {
        System.out.println("+-----+------------+------------+------------+--------------+----------+------------+------------+");
        System.out.printf("| No. | %-10s | %-10s | %-10s | %-12s | %-8s | %-10s | %-10s |\n",
                titles.get(0), titles.get(1), titles.get(2), titles.get(3), titles.get(4), titles.get(5), titles.get(6));
        System.out.println("+-----+------------+------------+------------+--------------+----------+------------+------------+");
    }

    private static void printOrderDetail(OrderDetail orderDetail) {
        System.out.printf("| %02d  | %-10s | %-10s | %-10s | %-12s | %8s | %10s | %10s |\n", orderDetail.rowNum,
                orderDetail.getOrderDate(), orderDetail.getRegion(), orderDetail.getRep(),
                orderDetail.getItem(), orderDetail.getUnits(), orderDetail.getUnitCost(), orderDetail.getTotal());
    }

    @lombok.Data
    @WorkbookRecord(titleRow = 0, endSheet = 1, startRow = 1)
    public static class OrderDetail {
        @Metadata(MetadataType.ROW_NUMBER)
        private Integer rowNum;

        @Property(column = 0)
        private LocalDate orderDate;

        @Property(column = 1)
        private String region;

        @Property(column = 2)
        private String rep;

        @Property(column = 3)
        private String item;

        @Property(column = 4)
        private BigInteger units;

        @Property(column = 5)
        private BigDecimal unitCost;

        @Property(column = 6)
        private BigDecimal total;
    }

}
```

#### Sample Output
```
+-----+------------+------------+------------+--------------+----------+------------+------------+
| No. | OrderDate  | Region     | Rep        | Item         | Units    | Unit Cost  | Total      |
+-----+------------+------------+------------+--------------+----------+------------+------------+
|  01 | 2021-01-06 | East       | Jones      | Pencil       |       95 |       1.99 |     189.05 |
|  02 | 2021-01-23 | Central    | Kivell     | Binder       |       50 |      19.99 |     999.50 |
|  03 | 2021-02-09 | Central    | Jardine    | Pencil       |       36 |       4.99 |     179.64 |
|  04 | 2021-02-26 | Central    | Gill       | Pen          |       27 |      19.99 |     539.73 |
|  05 | 2021-03-15 | West       | Sorvino    | Pencil       |       56 |       2.99 |     167.44 |
|  06 | 2021-04-01 | East       | Jones      | Binder       |       60 |       4.99 |     299.40 |
|  07 | 2021-04-18 | Central    | Andrews    | Pencil       |       75 |       1.99 |     149.25 |
...
```

### Legacy SAX Reading
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

#### Sample Output
```xml
<?xml version="1.0" encoding="UTF-8"?>
<document>
    <sheet index="0" name="Sheet1">
        <row index="0">
            <cell index="0" javaType="java.lang.String">Integer or Long</cell>
            <cell index="1" javaType="java.lang.Integer">1</cell>
            <cell index="2" javaType="java.lang.Integer">-10</cell>
            <cell index="3" javaType="java.lang.Long">10000000000</cell>
            <cell index="4" javaType="java.lang.Long">-10000000000</cell>
        </row>
        <row index="1">
            <cell index="0" javaType="java.lang.String">Double</cell>
            <cell index="1" javaType="java.lang.Double">9.9</cell>
            <cell index="2" javaType="java.lang.Double">-0.5555555556</cell>
            <cell index="3" javaType="java.lang.Integer">0</cell>
            <cell index="4" javaType="java.math.BigDecimal">0.0001234568</cell>
            <cell index="5" javaType="java.lang.Double">222.2222222</cell>
        </row>
        <row index="2">
            <cell index="0" javaType="java.lang.String">LocalTime</cell>
            <cell index="1" javaType="java.time.LocalTime">10:30</cell>
            <cell index="2" javaType="java.time.LocalTime">21:30:01</cell>
            <cell index="3" javaType="java.time.LocalTime">00:00</cell>
        </row>
    </sheet>
    <sheet index="1" name="Sheet2">
        <row index="0">
            <cell index="0" javaType="java.lang.String">Integer or Long</cell>
            <cell index="1" javaType="null"/>
            <cell index="2" javaType="null"/>
            <cell index="3" javaType="null"/>
            <cell index="4" javaType="null"/>
        </row>
    </sheet>
</document>
```
