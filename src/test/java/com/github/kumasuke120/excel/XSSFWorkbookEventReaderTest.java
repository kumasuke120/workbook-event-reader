package com.github.kumasuke120.excel;

import com.github.kumasuke120.util.ResourceUtil;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

class XSSFWorkbookEventReaderTest extends AbstractWorkbookEventReaderTest<XSSFWorkbookEventReader> {

    private static final String NORMAL_FILE_NAME = "workbook.xlsx";
    private static final String ENCRYPTED_FILE_NAME = "workbook-encrypted.xlsx";


    XSSFWorkbookEventReaderTest() {
        super(NORMAL_FILE_NAME, ENCRYPTED_FILE_NAME, XSSFWorkbookEventReader.class);
    }

    @Test
    @Override
    void read() {
        super.read();

        final Path filePath = ResourceUtil.getPathOfClasspathResource("ENGINES.xlsx");
        try (final WorkbookEventReader reader = pathConstructor().newInstance(filePath)) {
            reader.read(new WorkbookEventReader.EventHandler() {

                @Override
                public void onStartSheet(int sheetIndex, String sheetName) {
                    assertEquals("result 1", sheetName);
                }

                @Override
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
                    if (rowNum == 0 && columnNum == 0) {
                        assertEquals("ENGINE", cellValue.stringValue());
                    }
                }

            });
        }
    }

    @Test
    @Override
    void cancel() {
        super.cancel();

        dealWithReader(reader -> {
            assert reader instanceof XSSFWorkbookEventReader;

            final boolean[] cancelledRef = {false};
            final WorkbookEventReader.EventHandler handler = new WorkbookEventReader.EventHandler() {
                @Override
                public void onStartRow(int sheetIndex, int rowNum) {
                    reader.cancel();
                }

                @Override
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
                    throw new AssertionError();
                }

                @Override
                public void onReadCancelled() {
                    cancelledRef[0] = true;
                }
            };

            reader.read(handler);

            assertTrue(cancelledRef[0]);
        });
    }

    @Test
    @Override
    void close() {
        super.close();
    }

    @Test
    void setUse1904Windowing() {
        XSSFWorkbookEventReader.setUse1904Windowing(true);

        dealWithReader(reader -> {
            assert reader instanceof XSSFWorkbookEventReader;
            reader.read(new WorkbookEventReader.EventHandler() {
                @Override
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
                    if (sheetIndex == 0 && (rowNum == 3 || rowNum == 4) && columnNum == 1) {
                        if (!cellValue.isNull()) {
                            assertEquals(2022, cellValue.localDateValue().getYear());
                        }
                    }
                }
            });

            reader.close();
        });
    }

}
