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

public class AnnotationObjectReadPrinter {

    public static void main(String[] args) {
        final Path filePath = ResourceUtil.getPathOfClasspathResource("handler/sample-data.xlsx");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            // extracts the records to Java Objects
            final WorkbookRecordExtractor<OrderDetail> extractor = WorkbookRecordExtractor.ofRecord(OrderDetail.class);
            reader.read(extractor);

            // prints the result
            printResult(extractor);
        }
    }

    private static void printResult(WorkbookRecordExtractor<OrderDetail> extractor) {
        printOrderDetailTitle(extractor.getAllColumnTitles(0));
        for (OrderDetail orderDetail : extractor.getRecords(0)) {
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

    @SuppressWarnings("unused")
    @WorkbookRecord(titleRow = 0, endSheet = 1, startRow = 1)
    public static class OrderDetail {

        @Metadata(MetadataType.SHEET_NAME)
        private String lang;

        @Metadata(MetadataType.SHEET_INDEX)
        private Integer sheetIndex;

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

        public String getLang() {
            return lang;
        }

        public Integer getSheetIndex() {
            return sheetIndex;
        }

        public Integer getRowNum() {
            return rowNum;
        }

        public LocalDate getOrderDate() {
            return orderDate;
        }

        public String getRegion() {
            return region;
        }

        public String getRep() {
            return rep;
        }

        public String getItem() {
            return item;
        }

        public BigInteger getUnits() {
            return units;
        }

        public BigDecimal getUnitCost() {
            return unitCost;
        }

        public BigDecimal getTotal() {
            return total;
        }

        @Override
        public String toString() {
            return "OrderDetail[" + lang + ":" + sheetIndex + ":" + rowNum + "]{" +
                    "orderDate=" + orderDate +
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
