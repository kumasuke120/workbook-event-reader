# WorkbookEventReader
他の言語で表示: [English](README.md) | [中文](README_CN.md) | 日本語

イベントベースのワークブックリーダーです。[Apache POI](https://poi.apache.org/) Event API を再カプセル化して Excel ワークブックを処理し、
同じ統一されたインターフェイスで SpreadsheetML（.xlsx）とレガシー Excel ドキュメント（.xls）を簡単に読み取ることができます。
Excel ドキュメントから読み取ったすべての値は、対応する適切な Java データ型に変換できます。

## ブランチ
このリポジトリには 2 種類のブランチがあります。ブランチの種類ごとに、ブランチ名に独自のプレフィックスがあります：
- `java11-*`: Java 11 で記述され、多くの場合、新しい機能が追加されます (_デフォルトブランチ_)
- `java8-*`: Java 8 で記述され、より安定した機能を備えています (_プロダクション環境に推奨_)

## 要件
それぞれのブランチの要件は異なる場合があります：

| ブランチ             	| 同梱   	| `java11-*`       	| `java8-*`        	|
|--------------------	|---------	|------------------	|------------------	|
| Java Ver.          	| 無い      	| 11 onwards    	| 8 onwards     	|
| Apache POI Ver.    	| 4.0.0   	| 3.17 onwards  	| 3.17 onwards  	|
| Apache Xerces Ver. 	| 2.12.0  	| 2.0.0 onwards 	| 2.0.0 onwards 	|

_Apache POI 3.17 以降の要件の理由を確認するには、[ここ](https://bz.apache.org/bugzilla/show_bug.cgi?id=61034)をクリックしてください_

## ビルドとインストール
次のコマンドを使用してインストールすることができます：
```
$ ./mvnw clean install -DskipTests
```
または、ご自分の Maven 環境を使用してインストールすることもできます：
```
$ mvn clean install -DskipTests
```

`java11-*`ブランチの Maven 依存関係は次のとおりです：
```xml
<dependency>
    <groupId>app.kumasuke.excel</groupId>
    <artifactId>workbook-event-reader-java11</artifactId>
    <version>1.0.0</version>
</dependency>
```
`java8-*`ブランチの Maven 依存関係は次のとおりです：
```xml
<dependency>
    <groupId>app.kumasuke.excel</groupId>
    <artifactId>workbook-event-reader</artifactId>
    <version>1.0.0</version>
</dependency>
```

## サンプル
次のソースコードは、ワークブックを読み取り、そのコンテンツを整形式の XML 文書に変換します：
```java
public class ToXmlPrinter {
    public static void main(String[] args) {
        final Path filePath = Paths.get("workbook.xlsx");
        final XmlGenerator xmlGenerator = new XmlGenerator();
        try (final var reader = WorkbookEventReader.open(filePath)) {
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
            xml.append(/* 4つのスペース */"    ".repeat(currentIndentLevel));
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
_本例は Java 11 で記述されています。Java 8 を使用する場合は、一部を変更する必要がある場合があります_