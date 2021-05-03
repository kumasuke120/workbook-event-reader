package com.github.kumasuke120.excel;

import org.junit.jupiter.api.Test;

class CSVWorkbookEventReaderTest extends AbstractWorkbookEventReaderTest<CSVWorkbookEventReader>  {

    private static final String NORMAL_FILE_NAME = "ENGINES.csv";

    CSVWorkbookEventReaderTest() {
        super(NORMAL_FILE_NAME, "", CSVWorkbookEventReader.class);
        setSampleReadFileName("sample-output-csv.xml");
        setSampleCancelFileName("sample-output-csv.xml");
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