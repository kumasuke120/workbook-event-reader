package com.github.kumasuke120.excel;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * A {@link WorkbookEventReader} reads a csv workbook whose file extension
 * usually is <code>csv</code>
 */
public class CSVWorkbookEventReader extends AbstractWorkbookEventReader {

    private static final ThreadLocal<Charset> charsetLocal = ThreadLocal.withInitial(() -> StandardCharsets.UTF_8);
    private static final ThreadLocal<CSVFormat> formatLocal = ThreadLocal.withInitial(() -> CSVFormat.EXCEL);

    private CSVParser parser;
    private Charset charset;
    private CSVFormat format;

    /**
     * Creates a new {@link CSVWorkbookEventReader} based on the given file path.
     *
     * @param filePath the file path of the workbook
     */
    public CSVWorkbookEventReader(Path filePath) {
        super(filePath, null);
    }

    /**
     * Creates a new {@link CSVWorkbookEventReader} based on the given workbook {@link InputStream}.
     *
     * @param in {@link InputStream} of the workbook to be read
     */
    public CSVWorkbookEventReader(InputStream in) {
        super(in, null);
    }

    /**
     * Sets all following-opened instances of {@link CSVWorkbookEventReader} on the current thread using
     * the given {@link Charset} to read csv files.<br>
     *
     * @param charset character set to read csv files
     */
    public static void setCharset(Charset charset) {
        if (charset == null) {
            charsetLocal.remove();
        } else {
            charsetLocal.set(charset);
        }
    }

    /**
     * Sets all following-opened instances of {@link CSVWorkbookEventReader} on the current thread using
     * the given {@link CSVFormat} to read csv files.<br>
     *
     * @param format format to read csv files
     */
    public static void setCSVFormat(CSVFormat format) {
        if (format == null) {
            formatLocal.remove();
        } else {
            formatLocal.set(format);
        }
    }

    @Override
    void doOnStartOpen() {
        charset = charsetLocal.get();
        format = formatLocal.get();
    }

    @Override
    void doOpen(InputStream in, String password) throws Exception {
        final InputStreamReader reader = new InputStreamReader(in, charset);
        doOpen(reader);
    }

    @Override
    void doOpen(Path filePath, String password) throws Exception {
        final BufferedReader reader = Files.newBufferedReader(filePath, charset);
        doOpen(reader);
    }

    private void doOpen(Reader reader) throws IOException {
        parser = new CSVParser(reader, format);
    }

    @Override
    ReaderCleanAction createCleanAction() {
        return new CSVFReaderCleanAction(this);
    }

    @Override
    void doRead(EventHandler handler) {
        final EventHandler delegate = new CancelFastEventHandler(handler);

        /*
         * csv files don't have sheets, we triggers sheet-related events for compatibility;
         * treats the sheet-related event identical to document-related events
         */

        try {
            // handles onStartDocument
            delegate.onStartDocument();
            delegate.onStartSheet(0, "");

            int currentRowNumber = -1;
            for (final CSVRecord record : parser) {
                delegate.onStartRow(0, ++currentRowNumber);

                // handles cells
                int currentColumnNum;
                for (currentColumnNum = 0; currentColumnNum < record.size(); currentColumnNum++) {
                    final CellValue cellValue = getRecordCellValue(record, currentColumnNum);
                    delegate.onHandleCell(0, currentRowNumber, currentColumnNum, cellValue);
                }

                delegate.onEndRow(0, currentRowNumber);
            }

            // handles onEndDocument
            delegate.onEndSheet(0);
            delegate.onEndDocument();
        } catch (CancelReadingException e) {
            handler.onReadCancelled();
        }
    }

    private CellValue getRecordCellValue(CSVRecord record, int i) {
        /* gets and cleans value */
        String value = record.get(i);
        if (value == null || "".equals(value)) { // treats empty as null
            value = null;
        } else if (value.startsWith("\ufeff")) { // removes bom
            value = value.substring(1);
        }

        return CellValue.newInstance(value);
    }

    private static class CSVFReaderCleanAction extends ReaderCleanAction {
        private final CSVParser parser;

        CSVFReaderCleanAction(CSVWorkbookEventReader reader) {
            parser = reader.parser;
        }

        @Override
        void doClean() throws Exception {
            if (parser != null) {
                parser.close();
            }
        }
    }

    // stops EventHandler, not an actual exception
    private static class CancelReadingException extends RuntimeException {
    }

    // checks reading state before any event is triggered
    private class CancelFastEventHandler implements EventHandler {
        private final EventHandler handler;

        CancelFastEventHandler(EventHandler handler) {
            this.handler = handler;
        }

        @Override
        public void onStartDocument() {
            if (isReading()) {
                handler.onStartDocument();
            } else {
                throw new CancelReadingException();
            }
        }

        @Override
        public void onEndDocument() {
            if (isReading()) {
                handler.onEndDocument();
            } else {
                throw new CancelReadingException();
            }
        }

        @Override
        public void onStartSheet(int sheetIndex, String sheetName) {
            if (isReading()) {
                handler.onStartSheet(sheetIndex, sheetName);
            } else {
                throw new CancelReadingException();
            }
        }

        @Override
        public void onEndSheet(int sheetIndex) {
            if (isReading()) {
                handler.onEndSheet(sheetIndex);
            } else {
                throw new CancelReadingException();
            }
        }

        @Override
        public void onStartRow(int sheetIndex, int rowNum) {
            if (isReading()) {
                handler.onStartRow(sheetIndex, rowNum);
            } else {
                throw new CancelReadingException();
            }
        }

        @Override
        public void onEndRow(int sheetIndex, int rowNum) {
            if (isReading()) {
                handler.onEndRow(sheetIndex, rowNum);
            } else {
                throw new CancelReadingException();
            }
        }

        @Override
        public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
            if (isReading()) {
                handler.onHandleCell(sheetIndex, rowNum, columnNum, cellValue);
            } else {
                throw new CancelReadingException();
            }
        }

        @Override
        public void onReadCancelled() {
            throw new UnsupportedOperationException();
        }
    }

}
