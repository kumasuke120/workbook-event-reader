package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.excel.WorkbookProcessException;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Metadata;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Property;
import com.github.kumasuke120.excel.handler.WorkbookRecordExtractor.ExtractResult;
import com.github.kumasuke120.excel.util.CollectionUtils;
import com.github.kumasuke120.util.ResourceUtil;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class WorkbookRecordExtractorTest {

    private static final String TEST_RESOURCE_NAME = "handler/sample-data.xlsx";
    private static final String TEST_RESOURCE_NOTITLE_NAME = "handler/sample-data-notitle.xlsx";
    private static final String TEST_RESOURCE_WITHERROR_NAME = "handler/sample-data-witherror.xlsx";

    @Test
    void extract() {
        final Path filePath = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            final WorkbookRecordExtractor<Sheet0OrderDetail> extractor = WorkbookRecordExtractor.ofRecord(Sheet0OrderDetail.class);
            assertTrue(CollectionUtils.isEmpty(extractor.getAllRecords()));
            assertTrue(CollectionUtils.isEmpty(extractor.getRecords(0)));
            assertTrue(CollectionUtils.isEmpty(extractor.getRecords(1)));

            final List<Sheet0OrderDetail> result = extractor.extract(reader);
            assertEquals(extractor.getRecords(0), result);
            assertTrue(CollectionUtils.isEmpty(extractor.getRecords(1)));

            assertTrue(CollectionUtils.isEmpty(extractor.getAllColumnTitles(0)));
            assertEquals("", extractor.getColumnTitle(0, 0));

            assertEquals(42, result.size());
            assertEquals("eng", result.get(0).lang);
            assertEquals(0, result.get(0).sheetIndex);
            assertEquals(2, result.get(1).rowNum);
            assertEquals(LocalDate.of(2021, 2, 9), result.get(2).orderDate);
            assertEquals(LocalDate.of(2021, 2, 9), result.get(2).orderDate);
            assertEquals("Central", result.get(16).region);
            assertEquals("Parent", result.get(18).rep);
            assertEquals("Pen Set", result.get(21).item);
            assertEquals(80, result.get(30).units);
            assertEquals(new BigDecimal("4.99"), result.get(32).unitCost);
            assertEquals(new BigDecimal("1879.06"), result.get(41).total);
        }
    }

    @Test
    void extractWithTitle() {
        // reads titles from workbook row
        final Path filePath = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            final WorkbookRecordExtractor<Sheet0TitleRowOrderDetail> extractor = WorkbookRecordExtractor.ofRecord(Sheet0TitleRowOrderDetail.class);
            final List<Sheet0TitleRowOrderDetail> result = extractor.extract(reader);

            final List<String> columnTitles = extractor.getAllColumnTitles(0);
            assertTrue(CollectionUtils.isNotEmpty(columnTitles));
            assertEquals("OrderDate", columnTitles.get(0));
            assertEquals("Region", columnTitles.get(1));
            assertEquals("Rep", columnTitles.get(2));
            assertEquals("Item", columnTitles.get(3));
            assertEquals("Units", columnTitles.get(4));
            assertEquals("Unit Cost", columnTitles.get(5));
            assertEquals("Total", columnTitles.get(6));
            assertEquals("OrderDate", extractor.getColumnTitle(0, 0));
            assertEquals("Region", extractor.getColumnTitle(0, 1));
            assertEquals("Rep", extractor.getColumnTitle(0, 2));
            assertEquals("Item", extractor.getColumnTitle(0, 3));
            assertEquals("Units", extractor.getColumnTitle(0, 4));
            assertEquals("Unit Cost", extractor.getColumnTitle(0, 5));
            assertEquals("Total", extractor.getColumnTitle(0, 6));

            assertEquals("", extractor.getColumnTitle(1, 0));

            assertEquals(43, result.size());
        }

        // uses user-specified titles
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            final WorkbookRecordExtractor<Sheet1TitleRowOrderDetailWithTitle> extractor = WorkbookRecordExtractor.ofRecord(Sheet1TitleRowOrderDetailWithTitle.class);
            final List<Sheet1TitleRowOrderDetailWithTitle> result = extractor.extract(reader);

            final List<String> columnTitles = extractor.getAllColumnTitles(1);
            assertTrue(CollectionUtils.isNotEmpty(columnTitles));
            assertEquals("订单日期", columnTitles.get(0));
            assertEquals("地区", columnTitles.get(1));
            assertEquals("销售代表", columnTitles.get(2));
            assertEquals("商品", columnTitles.get(3));
            assertEquals("数量", columnTitles.get(4));
            assertEquals("单价（人民币）", columnTitles.get(5));
            assertEquals("总价（人民币）", columnTitles.get(6));
            assertEquals("订单日期", extractor.getColumnTitle(1, 0));
            assertEquals("地区", extractor.getColumnTitle(1, 1));
            assertEquals("销售代表", extractor.getColumnTitle(1, 2));
            assertEquals("商品", extractor.getColumnTitle(1, 3));
            assertEquals("数量", extractor.getColumnTitle(1, 4));
            assertEquals("单价（人民币）", extractor.getColumnTitle(1, 5));
            assertEquals("总价（人民币）", extractor.getColumnTitle(1, 6));

            assertEquals("", extractor.getColumnTitle(0, 0));

            assertEquals(43, result.size());
        }

        // workbook row has no titles
        final Path filePath2 = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NOTITLE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath2)) {
            final WorkbookRecordExtractor<Sheet0TitleRowOrderDetail> extractor = WorkbookRecordExtractor.ofRecord(Sheet0TitleRowOrderDetail.class);
            assertTrue(CollectionUtils.isEmpty(extractor.getAllColumnTitles(0)));
            assertTrue(CollectionUtils.isEmpty(extractor.getAllColumnTitles(1)));
            assertEquals("", extractor.getColumnTitle(0, 0));

            final List<Sheet0TitleRowOrderDetail> result = extractor.extract(reader);

            assertTrue(CollectionUtils.isEmpty(extractor.getAllColumnTitles(0)));
            assertTrue(CollectionUtils.isEmpty(extractor.getAllColumnTitles(1)));
            assertEquals("", extractor.getColumnTitle(0, 0));
            assertEquals(43, result.size());
        }
    }

    @Test
    void extractTwoSheets() {
        final Path filePath = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            final WorkbookRecordExtractor<SheetsOrderDetail> extractor = WorkbookRecordExtractor.ofRecord(SheetsOrderDetail.class);
            final List<SheetsOrderDetail> result = extractor.extract(reader);

            assertEquals(18, result.size());
            assertEquals("eng", result.get(0).lang);
            assertEquals(0, result.get(0).sheetIndex);
            assertEquals(2, result.get(1).rowNum);
            assertEquals("Binder", result.get(5).item);
            assertEquals("chs", result.get(9).lang);
            assertEquals(1, result.get(9).sheetIndex);
            assertEquals(2, result.get(10).rowNum);
            assertEquals("江苏", result.get(15).region);
        }
    }

    @Test
    void testWithError() {
        final Path filePath = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_WITHERROR_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            final WorkbookRecordExtractor<OrderDetailV2> extractor = WorkbookRecordExtractor.ofRecord(OrderDetailV2.class);
            extractor.extract(reader);
            final List<OrderDetailV2> records = extractor.getRecords(0);
            assertEquals(34, records.size());
            final List<ExtractResult<OrderDetailV2>> failExtractResults = extractor.getFailExtractResults(0);
            assertEquals(10, failExtractResults.size());
        }
    }

    @Test
    void constructorException() {
        final Path filePath = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            assertThrows(WorkbookRecordException.class, () -> WorkbookRecordExtractor.ofRecord(NoConstructorOrderDetail.class));

            final WorkbookRecordExtractor<ErrorConstructorOrderDetail> extractor2 = WorkbookRecordExtractor.ofRecord(ErrorConstructorOrderDetail.class);
            assertThrows(WorkbookProcessException.class, () -> extractor2.extract(reader));

            final WorkbookRecordExtractor<AbstractrOrderDetail> extractor3 = WorkbookRecordExtractor.ofRecord(AbstractrOrderDetail.class);
            assertThrows(WorkbookProcessException.class, () -> extractor3.extract(reader));

        }
    }

    @SuppressWarnings("unused")
    public static class OrderDetail {

        @Metadata(MetadataType.SHEET_NAME)
        String lang;

        @Metadata(MetadataType.SHEET_INDEX)
        Integer sheetIndex;

        @Metadata(MetadataType.ROW_NUMBER)
        Integer rowNum;

        @Property(column = 0)
        LocalDate orderDate;

        @Property(column = 1)
        String region;

        @Property(column = 2)
        String rep;

        @Property(column = 3)
        String item;

        @Property(column = 4)
        int units;

        @Property(column = 5)
        BigDecimal unitCost;

        @Property(column = 6)
        BigDecimal total;

    }

    @SuppressWarnings("unused")
    public static class OrderDetailWithTitle {

        @Property(column = 0, title = "订单日期")
        LocalDate orderDate;

        @Property(column = 1, title = "地区")
        String region;

        @Property(column = 2, title = "销售代表")
        String rep;

        @Property(column = 3, title = "商品")
        String item;

        @Property(column = 4, title = "数量")
        int units;

        @Property(column = 5)
        BigDecimal unitCost;

        @Property(column = 6)
        BigDecimal total;

    }

    @SuppressWarnings("unused")
    @WorkbookRecord(titleRow = 0, endSheet = 1, startRow = 1)
    public static class OrderDetailV2 {

        @Metadata(MetadataType.ROW_NUMBER)
        Integer rowNum;

        @Property(column = 0)
        LocalDate orderDate;

        @Property(column = 1, valueMethod = "handleRegion")
        Region region;

        @Property(column = 2)
        String rep;

        @Property(column = 3)
        String item;

        @Property(column = 4)
        int units;

        @Property(column = 5)
        BigDecimal unitCost;

        @Property(column = 6)
        BigDecimal total;

        public Region handleRegion(CellValue cellValue) {
            final String region = cellValue.trim().stringValue().toUpperCase();
            return Region.valueOf(region);
        }

    }

    public enum Region {
        CENTRAL, EAST, WEST
    }

    @WorkbookRecord(titleRow = 0, endSheet = 1, startRow = 1)
    public static class Sheet0TitleRowOrderDetail extends OrderDetail {
    }

    @WorkbookRecord(endSheet = 1, startRow = 1, endRow = 43)
    public static class Sheet0OrderDetail extends OrderDetail {
    }

    @WorkbookRecord(endSheet = 2, startRow = 1, endRow = 10)
    public static class SheetsOrderDetail extends OrderDetail {
    }

    @WorkbookRecord
    public static abstract class AbstractrOrderDetail extends OrderDetail {
    }

    @WorkbookRecord
    public static class NoConstructorOrderDetail extends OrderDetail {
        private NoConstructorOrderDetail() {
        }
    }

    @WorkbookRecord
    public static class ErrorConstructorOrderDetail extends OrderDetail {
        public ErrorConstructorOrderDetail() {
            throw new UnsupportedOperationException();
        }
    }

    @WorkbookRecord(titleRow = 0, startSheet = 1, endSheet = 2, startRow = 1)
    public static class Sheet1TitleRowOrderDetailWithTitle extends OrderDetailWithTitle {
    }

}