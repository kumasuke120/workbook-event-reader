package com.github.kumasuke120.demo;

import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.excel.handler.WorkbookRecord;
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

        @WorkbookRecord.SheetIndex
        private Integer sheetIndex;

        @WorkbookRecord.RowNumber
        private Integer rowNum;

        @WorkbookRecord.Property(column = 0)
        private LocalDate orderDate;

        @WorkbookRecord.Property(column = 1)
        private String region;

        @WorkbookRecord.Property(column = 2)
        private String rep;

        @WorkbookRecord.Property(column = 3)
        private String item;

        @WorkbookRecord.Property(column = 4)
        private BigInteger units;

        @WorkbookRecord.Property(column = 5)
        private BigDecimal unitCost;

        @WorkbookRecord.Property(column = 6)
        private BigDecimal total;

        @Override
        public String toString() {
            return "OfficeSupplySalesData{" +
                    "sheetIndex=" + sheetIndex +
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
