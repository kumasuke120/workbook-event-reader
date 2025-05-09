# WorkbookEventReader
查看其他语言: [English](README.md) | 中文

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![Build Status](https://github.com/kumasuke120/workbook-event-reader/actions/workflows/build.yml/badge.svg)](https://github.com/kumasuke120/workbook-event-reader/actions/workflows/build.yml) [![codecov](https://codecov.io/gh/kumasuke120/workbook-event-reader/branch/master/graph/badge.svg)](https://codecov.io/gh/kumasuke120/workbook-event-reader)

## 特色
- 使用易用的注解读取 Excel 文件，并将其转换为 Java 对象
- 事件流式的工作簿阅读器，提供统一接口处理 SpreadsheetML(.xlsx) 、传统 Excel 文档(.xls) 和 CSV 文件
- 高效率读取大工作簿文件，内存占用低
- 单文件可多次读取；单元格值读取简单，可轻松类型转换

## 支持文档格式
- *.xlsx
- *.xls
- *.csv (支持编码自动检测)

## 环境需求
本项目依赖的外部库如下所示：

| 依赖                    	   | 内置   	    | 最低		                                                                                                     |
|---------------------------|-----------|----------------------------------------------------------------------------------------------------------|
| Java 版本             	     | 无      	  | 8   	                                                                                                    |
| Apache POI 版本       	     | 4.1.2   	 | 3.17<sup><a href="https://bz.apache.org/bugzilla/show_bug.cgi?id=61034" target="_blank">[?]</a></sup>  	 |
| Apache Xerces 版本        	 | 2.12.2  	 | 2.0.0                                                                                                    |
| Apache Commons CSV 版本 	   | 1.8     	 | 1.0                                                                                                      |

## 构建与安装
可以使用以下命令安装：
```
$ ./mvnw clean install -DskipTests
```
或者可以选用自己的 Maven：
```
$ mvn clean install -DskipTests
```

本项目的 Maven 依赖为:
```xml
<dependency>
    <groupId>com.github.kumasuke120</groupId>
    <artifactId>workbook-event-reader</artifactId>
    <version>1.6.3</version>
</dependency>
```

## 快速上手
### 注解方式读取
借助注解，您可以使用以下代码将工作簿内容转换为 Java 对象并打印其内容：

```java
public class AnnotationObjectReadPrinter {

    public static void main(String[] args) {
        final Path filePath = Paths.get("workbook.xlsx");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            // 提取记录转换为 Java 对象
            final WorkbookRecordExtractor<OrderDetail> extractor = WorkbookRecordExtractor.ofRecord(OrderDetail.class);
            reader.read(extractor);

            // 打印结果
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
#### 样例输出结果
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

### 传统 SAX 方式读取
以下的代码读取工作簿，并将它的内容转换为格式化后的 XML 文档：
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
                sb.append(/* 4 个空格 */"    ");
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

#### 样例输出结果
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
