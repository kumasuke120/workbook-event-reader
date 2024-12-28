package com.github.kumasuke120.demo;

import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.excel.handler.WorkbookRecord;
import com.github.kumasuke120.excel.handler.WorkbookRecord.CellValueType;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Metadata;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Property;
import com.github.kumasuke120.excel.handler.WorkbookRecordExtractor;
import com.github.kumasuke120.util.ResourceUtil;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;

public class SampleDataExtractor {

    public static void main(String[] args) {
        final Path filePath = ResourceUtil.getPathOfClasspathResource("handler/SampleData.xlsx");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            List<OfficeSupplySalesData> result = WorkbookRecordExtractor.extract(reader, OfficeSupplySalesData.class);
            result.forEach(System.out::println);
        }
    }

    @SuppressWarnings("unused")
    @WorkbookRecord(startSheet = 1, endSheet = 2, startRow = 1, endColumn = 7)
    public static class OfficeSupplySalesData {

        @Metadata(MetadataType.SHEET_INDEX)
        private Integer sheetIndex;

        @Metadata(MetadataType.SHEET_NAME)
        private String sheetName;

        @Metadata(MetadataType.ROW_NUMBER)
        private Integer rowNum;

        @Property(column = 0, valueType = CellValueType.DATE)
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

        @Override
        public String toString() {
            return "OfficeSupplySalesData{" +
                    "sheetIndex=" + sheetIndex +
                    ", sheetName='" + sheetName + '\'' +
                    ", rowNum=" + rowNum +
                    ", orderDate=" + orderDate +
                    ", region='" + region + '\'' +
                    ", rep='" + rep + '\'' +
                    ", item='" + item + '\'' +
                    ", units=" + units +
                    ", unitCost=" + unitCost +
                    ", total=" + total +
                    '}';
        }
    }

}
