package com.github.kumasuke120.excel;

import com.github.kumasuke120.util.ResourceUtil;
import org.apache.commons.csv.CSVFormat;
import org.jetbrains.annotations.NotNull;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;

class CSVWorkbookEventReaderTest extends AbstractWorkbookEventReaderTest<CSVWorkbookEventReader> {

    private static final String NORMAL_FILE_NAME = "ENGINES.csv";

    CSVWorkbookEventReaderTest() {
        super(NORMAL_FILE_NAME, "", CSVWorkbookEventReader.class);
        setSampleReadFileName("sample-output-csv.xml");
        setSampleCancelFileName("sample-output-csv.xml");
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

    @Test
    void setCharset() {
        CSVWorkbookEventReader.setCharset(StandardCharsets.UTF_16BE);
        CSVWorkbookEventReader.setCSVFormat(CSVFormat.DEFAULT);

        try {
            final Path filePath = ResourceUtil.getPathOfClasspathResource("workbook-utf16be.csv");
            try (final WorkbookEventReader reader = pathConstructor().newInstance(filePath)) {
                reader.read(new WorkbookEventReader.EventHandler() {
                    @Override
                    public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
                        if (rowNum == 0 && columnNum == 0) {
                            assertEquals("中文", cellValue.originalValue());
                        } else if (rowNum == 0 && columnNum == 1) {
                            assertEquals("UTF16BE", cellValue.originalValue());
                        }
                    }
                });
            }
        } finally {
            CSVWorkbookEventReader.setCharset(null);
            CSVWorkbookEventReader.setCSVFormat(null);
        }
    }

}