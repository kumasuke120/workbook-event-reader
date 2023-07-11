package com.github.kumasuke120.excel;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.util.IOUtils;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * A {@link WorkbookEventReader} reads a csv workbook whose file extension
 * usually is <code>csv</code>
 */
@SuppressWarnings("unused")
public class CSVWorkbookEventReader extends AbstractWorkbookEventReader {

    private static final ThreadLocal<Charset> charsetLocal = new ThreadLocal<>();
    private static final ThreadLocal<CSVFormat> formatLocal = ThreadLocal.withInitial(() -> CSVFormat.EXCEL);

    private byte[] content;
    private Charset charset;
    private CSVFormat format;

    /**
     * Creates a new {@link CSVWorkbookEventReader} based on the given file path.
     *
     * @param filePath the file path of the workbook
     */
    public CSVWorkbookEventReader(@NotNull(exception = NullPointerException.class) Path filePath) {
        super(filePath, null);
    }

    /**
     * Creates a new {@link CSVWorkbookEventReader} based on the given workbook {@link InputStream}.
     *
     * @param in {@link InputStream} of the workbook to be read
     */
    public CSVWorkbookEventReader(@NotNull(exception = NullPointerException.class) InputStream in) {
        super(in, null);
    }

    /**
     * Sets all following-opened instances of {@link CSVWorkbookEventReader} on the current thread using
     * the given {@link Charset} to read csv files.<br>
     *
     * @param charset character set to read csv files
     */
    public static void setCharset(@Nullable Charset charset) {
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
    public static void setCSVFormat(@Nullable CSVFormat format) {
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
    void doOpen(@NotNull InputStream in, @Nullable String password) throws Exception {
        Exception thrown = null;
        try {
            content = IOUtils.toByteArray(in);
        } catch (Exception e) {
            thrown = e;
        } finally {
            suppressClose(in, thrown);
        }
    }

    @Override
    void doOpen(@NotNull Path filePath, @Nullable String password) throws Exception {
        final InputStream in = Files.newInputStream(filePath);
        doOpen(in, password);
    }

    @Override
    @NotNull
    ReaderCleanAction createCleanAction() {
        return new CSVFReaderCleanAction(this);
    }

    @Override
    void doRead(@NotNull EventHandler handler) throws Exception {
        // creates a new instance of Parser everytime to enable passes of readings
        try (final CSVParser parser = createParser()) {

            /*
             * csv files don't have sheets, we trigger sheet-related events for compatibility;
             * treats the sheet-related event identical to document-related events
             */

            // handles onStartDocument
            handler.onStartDocument();
            handler.onStartSheet(0, "");

            int currentRowNumber = -1;
            for (final CSVRecord record : parser) {
                handler.onStartRow(0, ++currentRowNumber);

                // handles cells
                for (int currentColumnNum = 0; currentColumnNum < record.size(); currentColumnNum++) {
                    final CellValue cellValue = getRecordCellValue(record, currentColumnNum);
                    handler.onHandleCell(0, currentRowNumber, currentColumnNum, cellValue);
                }

                handler.onEndRow(0, currentRowNumber);
            }

            // handles onEndDocument
            handler.onEndSheet(0);
            handler.onEndDocument();
        }
    }

    private CSVParser createParser() throws IOException {
        final ByteArrayInputStream in = new ByteArrayInputStream(content);
        final InputStreamReader reader = new InputStreamReader(in, getCharset());
        return new CSVParser(reader, format);
    }

    private Charset getCharset() {
        if (charset == null) {
            final ByteArrayCharsetDetector detector = new ByteArrayCharsetDetector(content);
            charset = detector.detect();

            if (charset == null) {
                charset = StandardCharsets.UTF_8;
            }
        }

        return charset;
    }

    @NotNull
    private CellValue getRecordCellValue(@NotNull CSVRecord record, int i) {
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
        CSVFReaderCleanAction(@NotNull CSVWorkbookEventReader reader) {
        }

        @Override
        void doClean() {
            // no-op
        }
    }

}
