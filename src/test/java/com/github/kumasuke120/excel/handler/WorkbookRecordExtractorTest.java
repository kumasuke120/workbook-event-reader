package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.excel.WorkbookProcessException;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Metadata;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.handler.WorkbookRecord.Property;
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

    @Test
    void extract() {
        final Path filePath = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            final WorkbookRecordExtractor<Sheet0OrderDetail> extractor = WorkbookRecordExtractor.ofRecord(Sheet0OrderDetail.class);
            final List<Sheet0OrderDetail> result = extractor.extract(reader);

            assertTrue(CollectionUtils.isEmpty(extractor.getColumnTitles()));
            assertEquals("", extractor.getColumnTitle(0));

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
        final Path filePath = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            final WorkbookRecordExtractor<Sheet0WithTitleOrderDetail> extractor = WorkbookRecordExtractor.ofRecord(Sheet0WithTitleOrderDetail.class);
            final List<Sheet0WithTitleOrderDetail> result = extractor.extract(reader);

            final List<String> columnTitles = extractor.getColumnTitles();
            assertTrue(CollectionUtils.isNotEmpty(columnTitles));
            assertEquals("OrderDate", columnTitles.get(0));
            assertEquals("Region", columnTitles.get(1));
            assertEquals("Rep", columnTitles.get(2));
            assertEquals("Item", columnTitles.get(3));
            assertEquals("Units", columnTitles.get(4));
            assertEquals("Unit Cost", columnTitles.get(5));
            assertEquals("Total", columnTitles.get(6));
            assertEquals("OrderDate", extractor.getColumnTitle(0));
            assertEquals("Region", extractor.getColumnTitle(1));
            assertEquals("Rep", extractor.getColumnTitle(2));
            assertEquals("Item", extractor.getColumnTitle(3));
            assertEquals("Units", extractor.getColumnTitle(4));
            assertEquals("Unit Cost", extractor.getColumnTitle(5));
            assertEquals("Total", extractor.getColumnTitle(6));

            assertEquals(43, result.size());
        }

        final Path filePath2 = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NOTITLE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath2)) {
            final WorkbookRecordExtractor<Sheet0WithTitleOrderDetail> extractor = WorkbookRecordExtractor.ofRecord(Sheet0WithTitleOrderDetail.class);
            assertNull(extractor.getColumnTitles());
            assertNull(extractor.getColumnTitles(1));
            assertNull(extractor.getColumnTitle(0));

            final List<Sheet0WithTitleOrderDetail> result = extractor.extract(reader);

            assertTrue(CollectionUtils.isEmpty(extractor.getColumnTitles()));
            assertTrue(CollectionUtils.isEmpty(extractor.getColumnTitles(1)));
            assertEquals("", extractor.getColumnTitle(0));
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
    void constructorException() {
        final Path filePath = ResourceUtil.getPathOfClasspathResource(TEST_RESOURCE_NAME);
        try (final WorkbookEventReader reader = WorkbookEventReader.open(filePath)) {
            final WorkbookRecordExtractor<NoConstructorOrderDetail> extractor = WorkbookRecordExtractor.ofRecord(NoConstructorOrderDetail.class);
            assertThrows(WorkbookProcessException.class, () -> extractor.extract(reader));

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

    @WorkbookRecord(titleRow = 0, endSheet = 1, startRow = 1)
    public static class Sheet0WithTitleOrderDetail extends OrderDetail {
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

}