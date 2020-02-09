package app.kumasuke.excel;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.util.IOUtils;

import java.io.ByteArrayInputStream;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.lang.ref.Cleaner;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.*;
import java.util.Date;
import java.util.Map;
import java.util.Objects;
import java.util.TimeZone;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * The base class for {@link WorkbookEventReader}, containing common methods and utilities
 */
abstract class AbstractWorkbookEventReader implements WorkbookEventReader {

    private static final Cleaner cleaner = Cleaner.create();

    private final Cleaner.Cleanable cleanable;

    private volatile boolean closed = false;
    private volatile boolean reading = false;

    /**
     * Creates a new {@link AbstractWorkbookEventReader} based on the given file {@link InputStream}
     * and the given password if possible.
     *
     * @param in       {@link InputStream} of the workbook to be read
     * @param password password to open the file
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    AbstractWorkbookEventReader(InputStream in, String password) {
        Objects.requireNonNull(in);

        try {
            doOpen(in, password);
        } catch (Exception e) {
            throw new WorkbookIOException("Cannot open workbook", e);
        }

        final ReaderCleanAction action = createCleanAction();
        cleanable = cleaner.register(this, action);
    }

    /**
     * Opens an {@link InputStream} from the given file path.
     *
     * @param filePath file path of the workbook
     * @return {@link InputStream} opened
     */
    static InputStream getWorkbookInputStream(Path filePath) {
        try {
            return Files.newInputStream(Objects.requireNonNull(filePath));
        } catch (IOException e) {
            throw new WorkbookIOException("Cannot open workbook", e);
        }
    }

    /**
     * Closes the {@link Closeable}. Exception thrown during closing will be suppressed and
     * add to the previous caught {@link Exception} if necessary. Otherwise the exception will
     * be thrown.
     *
     * @param closeable object to be closed
     * @param caught    previous caught {@link Exception}
     * @throws Exception previous caught {@link Exception} or {@link Exception} thrown when closing
     */
    static void suppressClose(Closeable closeable, Exception caught) throws Exception {
        Exception thrown = caught;
        try {
            closeable.close();
        } catch (IOException e) {
            if (thrown != null) {
                e.addSuppressed(thrown);
            }
            thrown = e;
        }

        if (thrown != null) throw thrown;
    }

    /**
     * Opens the specified {@link InputStream} with an appropriate {@link WorkbookEventReader} if possible.<br>
     * It opens the specified file with {@link XSSFWorkbookEventReader} and {@link HSSFWorkbookEventReader} in
     * a calculated order. If both creations fails, it will throw {@link WorkbookIOException}
     *
     * @param in       {@link InputStream} of the workbook to be opened
     * @param password password to open the file
     * @return {@link WorkbookEventReader} to read the specified file
     * @throws NullPointerException <code>in</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    static WorkbookEventReader autoOpen(InputStream in, String password, boolean firstTryXSSF) {
        Objects.requireNonNull(in);

        final byte[] bytes;
        try {
            bytes = IOUtils.toByteArray(in);
        } catch (IOException e) {
            throw new WorkbookIOException("Cannot open and read workbook", e);
        }

        final FileMagic magic = checkFileMagic(bytes);
        if (magic == FileMagic.OOXML) {
            return new XSSFWorkbookEventReader(new ByteArrayInputStream(bytes));
        } else {
            if (firstTryXSSF) {
                return openByOrder(() -> new XSSFWorkbookEventReader(new ByteArrayInputStream(bytes), password),
                                   () -> new HSSFWorkbookEventReader(new ByteArrayInputStream(bytes), password));
            } else {
                return openByOrder(() -> new HSSFWorkbookEventReader(new ByteArrayInputStream(bytes), password),
                                   () -> new XSSFWorkbookEventReader(new ByteArrayInputStream(bytes), password));
            }
        }
    }

    private static FileMagic checkFileMagic(byte[] bytes) {
        try (final var in = new ByteArrayInputStream(bytes);
             final InputStream stream = FileMagic.prepareToCheckMagic(in)) {
            return FileMagic.valueOf(stream);
        } catch (IOException e) {
            return null;
        }
    }

    private static WorkbookEventReader openByOrder(Supplier<WorkbookEventReader> firstConstructor,
                                                   Supplier<WorkbookEventReader> secondConstructor) {
        try {
            return firstConstructor.get();
        } catch (WorkbookIOException e1) {
            // EncryptedDocumentException means password incorrect, it should be thrown directly
            if (e1.getCause() instanceof EncryptedDocumentException) {
                throw e1;
            } else {
                try {
                    return secondConstructor.get();
                } catch (WorkbookIOException e2) {
                    if (!(e2.getCause() instanceof EncryptedDocumentException)) {
                        e2.addSuppressed(e1);
                    }
                    throw e2;
                }
            }
        }
    }

    /**
     * Opens the specified file with this {@link AbstractWorkbookEventReader}.<br>
     * <br>
     * * This method should not throw any {@link WorkbookEventReaderException}, because it is the duty of its
     * caller to wrap every exception it throws into {@link WorkbookEventReaderException}.
     *
     * @param in       {@link InputStream} of the workbook to be opened, and it won't be <code>null</code>
     * @param password password to open the file, and it might be <code>null</code>
     * @throws Exception any exception occurred during opening process
     */
    abstract void doOpen(InputStream in, String password) throws Exception;

    /**
     * Creates a resource-cleaning action to close all resources this {@link WorkbookEventReader}
     * has opened.<br>
     * <br>
     * * The implementation of {@link ReaderCleanAction} should not be an anonymous or non-static inner class.
     * In addition, it should not contain any reference of the containing {@link WorkbookEventReader}.
     *
     * @return a non-anonymous instance of {@link ReaderCleanAction}
     */
    abstract ReaderCleanAction createCleanAction();

    /**
     * {@inheritDoc}
     */
    @Override
    public final void read(EventHandler handler) {
        assertNotClosed();
        assertNotBeingRead();

        Objects.requireNonNull(handler);

        reading = true;
        try {
            doRead(handler);
        } catch (Exception e) {
            if (e instanceof WorkbookEventReaderException) {
                throw (WorkbookEventReaderException) e;
            } else {
                throw new WorkbookProcessException(e);
            }
        } finally {
            reading = false;
        }
    }

    /**
     * Starts to read the workbook through event handling, triggering events on the {@link EventHandler}
     * in a reasonable and recursive order: Document, Sheet, Row and Cell.<br>
     * <br>
     * * This method should be able to called multiple times as long as this {@link AbstractWorkbookEventReader}
     * is not closed.<br>
     * * This method or its user may throw a {@link WorkbookEventReaderException} which will be re-thrown in
     * {@link #read(EventHandler)}.
     *
     * @param handler an non-<code>null</code> {@link EventHandler} that handles read events as reading process going
     * @throws Exception any exception occurred during reading process
     */
    abstract void doRead(EventHandler handler) throws Exception;

    /**
     * {@inheritDoc}
     */
    @Override
    public final void cancel() {
        assertNotClosed();
        assertBeingRead();

        reading = false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public final void close() {
        if (!closed) {
            reading = false;
            closed = true;

            cleanable.clean();
        }
    }

    /**
     * Returns current reading state of the reader.
     *
     * @return <code>true</code> if the reader is in reading state, otherwise <code>false</code>
     */
    boolean isReading() {
        return reading;
    }

    /**
     * Asserts the reader is not being closed. Otherwise it throws {@link IllegalReaderStateException}.
     *
     * @throws IllegalReaderStateException the reader has been closed
     */
    void assertNotClosed() {
        if (closed) {
            throw new IllegalReaderStateException("This '" + getClass().getSimpleName() + "' has been closed");
        }
    }

    /**
     * Asserts the reader is not being read. Otherwise it throws {@link IllegalReaderStateException}.
     *
     * @throws IllegalReaderStateException the reader is being read
     */
    void assertNotBeingRead() {
        if (reading) {
            throw new IllegalReaderStateException("This '" + getClass().getSimpleName() + "' is being read");
        }
    }

    /**
     * Asserts the reader is being read. Otherwise it throws {@link IllegalReaderStateException}.
     *
     * @throws IllegalReaderStateException the reader is not being read
     */
    private void assertBeingRead() {
        if (!reading) {
            throw new IllegalReaderStateException("This '" + getClass().getSimpleName() + "' is not being read");
        }
    }

    /**
     * An utility class that contains various methods for dealing with workbook
     */
    static class Util {
        private static final double MAX_EXCEL_DATE_EXCLUSIVE = 2958466;
        private static final Pattern cellReferencePattern = Pattern.compile("([A-Z]+)(\\d+)");

        private Util() {
            throw new UnsupportedOperationException();
        }

        /**
         * Tests if the given value is a whole number.
         *
         * @param value <code>double</code> value to be tested
         * @return <code>true</code> if the given value is a whole number, otherwise <code>false</code>
         */
        static boolean isAWholeNumber(double value) {
            return value % 1 == 0;
        }

        /**
         * Tests if the given value is a whole number that Java could represent with primitive type.
         *
         * @param value <code>String</code> value to be tested
         * @return <code>true</code> if the given value is a whole number, otherwise <code>false</code>
         */
        static boolean isAWholeNumber(String value) {
            if (value == null) return false;

            try {
                Long.parseLong(value);
                return true;
            } catch (NumberFormatException e) {
                return false;
            }
        }

        /**
         * Tests if the given value is a decimal fraction that Java could represent with primitive type.
         *
         * @param value <code>String</code> value to be tested
         * @return <code>true</code> if the given value is a decimal fraction, otherwise <code>false</code>
         */
        static boolean isADecimalFraction(String value) {
            if (value == null) return false;

            try {
                Double.parseDouble(value);
                return true;
            } catch (NumberFormatException e) {
                return false;
            }
        }

        /**
         * Parses the <code>String</code> value to <code>int</code> silently.<br>
         * It will return <code>defaultValue</code> if the <code>String</code> value could not be parsed.
         *
         * @param value the <code>String</code> value to be parsed
         * @return parsed <code>int</code> value if parse succeeds, otherwise the <code>defaultValue</code>
         */
        @SuppressWarnings("SameParameterValue")
        static int toInt(String value, int defaultValue) {
            if (value == null) return defaultValue;

            try {
                return Integer.parseInt(value);
            } catch (NumberFormatException e) {
                return defaultValue;
            }
        }

        /**
         * Tests if the given value is a valid excel date.
         *
         * @param excelDateValue excel date value
         * @return code>true</code> if the given value is a valid excel date, otherwise <code>false</code>
         */
        static boolean isValidExcelDate(double excelDateValue) {
            return DateUtil.isValidExcelDate(excelDateValue) &&
                    excelDateValue > 0 && excelDateValue < MAX_EXCEL_DATE_EXCLUSIVE;
        }

        /**
         * Converts the given excel date value to @{@link LocalTime}, {@link LocalDateTime} or
         * {@link LocalDate} accordingly.
         *
         * @param excelDateValue   excel date value
         * @param use1904Windowing <code>true</code> if date uses 1904 windowing,
         *                         or <code>false</code> if using 1900 date windowing.
         * @return converted @{@link LocalTime}, {@link LocalDateTime} or {@link LocalDate}
         */
        static Object toJsr310DateOrTime(double excelDateValue, boolean use1904Windowing) {
            if (isValidExcelDate(excelDateValue)) {
                final Date date = DateUtil.getJavaDate(excelDateValue, use1904Windowing,
                                                       TimeZone.getTimeZone("UTC"));
                final LocalDateTime localDateTime = toLocalDateTimeOffsetByUTC(date);

                if (Util.isAWholeNumber(excelDateValue)) { // date only
                    return localDateTime.toLocalDate();
                } else if (excelDateValue < 1) { // time only
                    return localDateTime.toLocalTime();
                } else { // date with time
                    return localDateTime;
                }
            } else {
                return null;
            }
        }

        private static LocalDateTime toLocalDateTimeOffsetByUTC(Date date) {
            final Instant instant = date.toInstant();
            return LocalDateTime.ofInstant(instant, ZoneId.of("UTC"));
        }

        /**
         * Converts given value to relative type.<br>
         * For instance, if a <code>double</code> could be treated as a <code>int</code>, it will be converted
         * to its <code>int</code> counterpart.
         *
         * @param value value to be converted
         * @return value converted accordingly
         */
        static Object toRelativeType(Object value) {
            if (value instanceof Double) {
                final var doubleValue = (double) value;
                if (Util.isAWholeNumber(doubleValue)) {
                    if (doubleValue > Integer.MAX_VALUE || doubleValue < Integer.MIN_VALUE) {
                        return (long) doubleValue;
                    } else {
                        return (int) doubleValue;
                    }
                }
            } else if (value instanceof Long) {
                final var longValue = (long) value;
                if (longValue >= Integer.MIN_VALUE && longValue <= Integer.MAX_VALUE) {
                    return (int) longValue;
                }
            }

            return value;
        }

        /**
         * Converts excel cell reference (e.g. A1) to a {@link Map.Entry} containing row number (starts with 0)
         * and column number (starts with 0).<br>
         * It will return <code>null</code> if either of them cannot be parse correctly.
         *
         * @param cellReference excel cell reference
         * @return row number and column number, both starts with 0 or <code>null</code>
         */
        static Map.Entry<Integer, Integer> cellReferenceToRowAndColumn(String cellReference) {
            assert cellReference != null;

            final Matcher cellRefMatcher = cellReferencePattern.matcher(cellReference);
            if (cellRefMatcher.matches()) {
                final String rawColumn = cellRefMatcher.group(1);
                final String rawRow = cellRefMatcher.group(2);

                final int rowNum = Integer.parseInt(rawRow) - 1;
                final int columnNum = columnNameToInt(rawColumn);

                return Map.entry(rowNum, columnNum);
            } else {
                return null;
            }
        }

        // converts 'A' to 0, 'B' to 1, ..., 'AA' to 26, and etc.
        private static int columnNameToInt(String columnName) {
            int index = 0;

            for (int i = 0; i < columnName.length(); i++) {
                char c = columnName.charAt(i);
                if (i == columnName.length() - 1) {
                    index = index * 26 + (c - 'A');
                } else {
                    index = index * 26 + (c - 'A' + 1);
                }
            }

            return index;
        }

        /**
         * Tests if the given index of format or format string stands for a text format.
         *
         * @param formatIndex  index of format
         * @param formatString format string
         * @return <code>true</code> if the given arguments stand for a text format, otherwise <code>false</code>
         */
        static boolean isATextFormat(int formatIndex, String formatString) {
            return formatIndex == 0x31 ||
                    (formatString != null && BuiltinFormats.getBuiltinFormat(formatString) == 0x31);
        }
    }

    /**
     * A clean action for closes the resources opened by a {@link WorkbookEventReader}
     */
    static abstract class ReaderCleanAction implements Runnable {
        @Override
        public final void run() {
            try {
                doClean();
            } catch (Exception e) {
                throw new WorkbookIOException("Exception encountered when closing the workbook file", e);
            }
        }

        /**
         * Closes the resources from the related {@link WorkbookEventReader}. <br>
         * <br>
         * * This method should not throw any {@link WorkbookEventReaderException}, because it is the duty of its
         * caller to wrap every exception it throws into {@link WorkbookEventReaderException}.
         *
         * @throws Exception any exception occurred during closing process
         */
        abstract void doClean() throws Exception;
    }

}
