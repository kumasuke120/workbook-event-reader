package com.github.kumasuke120.excel;

import com.github.kumasuke120.util.ResourceUtil;
import org.jetbrains.annotations.NotNull;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class XSSF1904WorkbookEventReaderTest extends AbstractWorkbookEventReaderTest<XSSFWorkbookEventReader> {

    private static final String NORMAL_FILE_NAME = "workbook-1904.xlsx";
    private static final String ENCRYPTED_FILE_NAME = "workbook-encrypted-1904.xlsx";

    XSSF1904WorkbookEventReaderTest() {
        super(NORMAL_FILE_NAME, ENCRYPTED_FILE_NAME, XSSFWorkbookEventReader.class);
    }

    @Test
    @Override
    void open() throws IOException {
        super.open();
    }

    @Test
    @Override
    void read() {
        super.read();

        final Path filePath = ResourceUtil.getPathOfClasspathResource("ENGINES.xlsx");
        try (final WorkbookEventReader reader = pathConstructor().newInstance(filePath)) {
            reader.read(new WorkbookEventReader.EventHandler() {
                @Override
                public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
                    assertEquals("result 1", sheetName);
                }

                @Override
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
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
    }

    @Test
    @Override
    void close() {
        super.close();
    }

}
