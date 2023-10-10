# WorkbookEventReader
查看其他语言: [English](README.md) | 中文

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![Build Status](https://github.com/kumasuke120/workbook-event-reader/actions/workflows/build.yml/badge.svg)](https://github.com/kumasuke120/workbook-event-reader/actions/workflows/build.yml) [![codecov](https://codecov.io/gh/kumasuke120/workbook-event-reader/branch/master/graph/badge.svg)](https://codecov.io/gh/kumasuke120/workbook-event-reader)

## 特色
- 事件流式的工作簿阅读器，内存占用小
- 统一接口处理 SpreadsheetML(.xlsx) 、传统 Excel 文档(.xls) 和 CSV 文件
- 同一文件可重复读取
- 单元格值读取简单，可轻松类型转换

## 支持文档格式
- *.xlsx
- *.xls
- *.csv (支持编码自动检测)

## 环境需求
本项目依赖的外部库如下所示：

| 依赖                    	   | 内置   	    | 最低		   |
|---------------------------|-----------|--------|
| Java 版本             	     | 无      	  | 8   	  |
| Apache POI 版本         	   | 4.1.2   	 | 3.17 	 |
| Apache Xerces 版本        	 | 2.12.2  	 | 2.0.0  |
| Apache Commons CSV 版本 	   | 1.8     	 | 1.0    |

_点击[这里](https://bz.apache.org/bugzilla/show_bug.cgi?id=61034)查看需要 Apache POI 3.17 或以上版本的原因_

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
    <version>1.3.3</version>
</dependency>
```

## 快速上手
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