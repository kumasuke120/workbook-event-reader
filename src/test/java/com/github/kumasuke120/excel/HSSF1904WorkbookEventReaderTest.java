package com.github.kumasuke120.excel;

import org.junit.jupiter.api.Test;

class HSSF1904WorkbookEventReaderTest extends AbstractWorkbookEventReaderTest<HSSFWorkbookEventReader> {

    private static final String NORMAL_FILE_NAME = "workbook-1904.xls";
    private static final String ENCRYPTED_FILE_NAME = "workbook-encrypted-1904.xls";

    HSSF1904WorkbookEventReaderTest() {
        super(NORMAL_FILE_NAME, ENCRYPTED_FILE_NAME, HSSFWorkbookEventReader.class);
    }

    @Test
    @Override
    void read() {
        super.read();
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
