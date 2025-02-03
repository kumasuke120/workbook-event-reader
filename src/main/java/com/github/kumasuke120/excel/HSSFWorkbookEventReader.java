package com.github.kumasuke120.excel;

import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

/**
 * A {@link WorkbookEventReader} reads a legacy workbook (Excel 97 - 2003) whose file extension
 * usually is <code>xls</code>
 */
@SuppressWarnings("unused")
public class HSSFWorkbookEventReader extends AbstractWorkbookEventReader {

    private static final short USER_CODE_CONTINUE = 0;
    private static final short USER_CODE_ABORT = Short.MIN_VALUE;

    private POIFSFileSystem fileSystem;
    private String password;
    private DataFormatter dataFormatter;

    /**
     * Creates a new {@link HSSFWorkbookEventReader} based on the given file path.
     *
     * @param filePath the file path of the workbook
     */
    public HSSFWorkbookEventReader(@NotNull(exception = NullPointerException.class) Path filePath) {
        this(filePath, null);
    }

    /**
     * Creates a new {@link HSSFWorkbookEventReader} based on the given file path using given password.
     *
     * @param filePath the file path of the workbook
     */
    public HSSFWorkbookEventReader(@NotNull(exception = NullPointerException.class) Path filePath,
                                   @Nullable String password) {
        super(filePath, password);
    }

    /**
     * Creates a new {@link HSSFWorkbookEventReader} based on the given workbook {@link InputStream}.
     *
     * @param in {@link InputStream} of the workbook to be read
     */
    public HSSFWorkbookEventReader(@NotNull(exception = NullPointerException.class) InputStream in) {
        this(in, null);
    }

    /**
     * Creates a new {@link HSSFWorkbookEventReader} based on the given workbook {@link InputStream}
     * and the given password if possible.
     *
     * @param in       {@link InputStream} of the workbook to be read
     * @param password password to open the file
     */
    public HSSFWorkbookEventReader(@NotNull(exception = NullPointerException.class) InputStream in,
                                   @Nullable String password) {
        super(in, password);
    }

    @Override
    void doOpen(@NotNull InputStream in, @Nullable String password) throws Exception {
        Exception thrown = null;
        try {
            fileSystem = new POIFSFileSystem(in); // consumes all stream data to memory
        } catch (Exception e) {
            thrown = e;
        } finally {
            suppressClose(in, thrown);
        }

        init();
        setWorkbookPassword(password);
    }

    @Override
    void doOpen(@NotNull Path filePath, @Nullable String password) throws Exception {
        final File file = filePath.toFile();
        fileSystem = new POIFSFileSystem(file, true);

        init();
        setWorkbookPassword(password);
    }

    private void init() {
        dataFormatter = new DataFormatter();
    }

    private void setWorkbookPassword(@Nullable String password) throws IOException {
        this.password = password; // records for later use
        checkWorkbookPassword(); // all documents should be checked
    }

    private void checkWorkbookPassword() throws IOException {
        final HSSFRequest request = new HSSFRequest();
        request.addListenerForAllRecords(new InstantAbortHSSFListener());

        processRequest(request);
    }

    @Override
    @NotNull
    ReaderCleanAction createCleanAction() {
        return new HSSFReaderCleanAction(this);
    }

    @Override
    void doRead(@NotNull EventHandler handler) throws Exception {
        handler.onStartDocument();

        final HSSFRequest request = new HSSFRequest();
        final ReaderHSSFListener readerListener = new ReaderHSSFListener(handler);
        request.addListenerForAllRecords(readerListener);

        // processes the document
        processRequest(request);

        handler.onEndDocument();
    }

    private void processRequest(@NotNull HSSFRequest request) throws IOException {
        try (final DocumentInputStream documentIs = fileSystem.createDocumentInputStream("Workbook")) {
            boolean passwordSet = false;
            String oldStoredUserPassword = null;
            try {
                if (password != null) { // set password if necessary
                    oldStoredUserPassword = Biff8EncryptionKey.getCurrentUserPassword();
                    Biff8EncryptionKey.setCurrentUserPassword(password);
                    passwordSet = true;
                }

                // processes the document
                final HSSFEventFactory factory = new HSSFEventFactory();
                factory.abortableProcessEvents(request, documentIs);
            } catch (HSSFUserException hue) {
                // cancelled by user-thrown exception, ReaderHSSFListener does not throw any
                // HSSFUserException or its subclasses
                // this exception won't trigger #onReadCancelled() event
            } finally {
                if (passwordSet) {
                    Biff8EncryptionKey.setCurrentUserPassword(oldStoredUserPassword);
                }
            }
        }
    }

    private static class HSSFReaderCleanAction extends ReaderCleanAction {
        private final POIFSFileSystem poifsFileSystem;

        HSSFReaderCleanAction(@NotNull HSSFWorkbookEventReader reader) {
            poifsFileSystem = reader.fileSystem;
        }

        @Override
        void doClean() throws Exception {
            if (poifsFileSystem != null) {
                poifsFileSystem.close();
            }
        }
    }

    // aborts reading at the first record
    private static class InstantAbortHSSFListener extends AbortableHSSFListener {
        @Override
        public short abortableProcessRecord(@NotNull Record record) {
            return USER_CODE_ABORT;
        }
    }

    private class ReaderHSSFListener extends AbortableHSSFListener {
        private final EventHandler handler;
        private final FormatTrackingHSSFListener formatTracker;

        private boolean use1904Windowing = false;
        private SSTRecord sharedStringTable;
        private Map<Integer, BoundSheetRecord> boundSheets;
        private Map<Integer, RowRecord> currentSheetRows;

        private int previousSheetIndex = -1;
        private int currentSheetIndex = -1;
        private int previousRowNumber = -1;
        private int currentRowNumber = -1;
        private int currentRowEndColumnNum = -1;

        private boolean previousSheetEndHandled = true;
        private boolean previousRowEndHandled = true;

        private Record previousRecord;

        private int tSheetIndex = -1;
        private int tRowNum = -1;

        private ReaderHSSFListener(@NotNull EventHandler handler) {
            this.handler = handler;
            this.formatTracker = new FormatTrackingHSSFListener(null);
        }

        @Override
        public short abortableProcessRecord(@NotNull Record record) {
            formatTracker.processRecordInternally(record); // records the formats and styles

            final short currentSid = record.getSid();
            switch (currentSid) {
                case BoundSheetRecord.sid: {
                    final BoundSheetRecord boundSheet = (BoundSheetRecord) record;
                    assert boundSheets != null;

                    boundSheets.put(++tSheetIndex, boundSheet);
                    break;
                }
                case EOFRecord.sid: {
                    // this record exists after header or sheet end
                    if (currentSheetIndex != -1) { // not header end
                        handleEndSheet(currentSheetIndex);
                    }

                    break;
                }
                case DateWindow1904Record.sid: {
                    final DateWindow1904Record dateWindow1904 = (DateWindow1904Record) record;
                    use1904Windowing = dateWindow1904.getWindowing() == 1;
                    break;
                }
                case SSTRecord.sid: {
                    sharedStringTable = (SSTRecord) record;
                    break;
                }
                case BOFRecord.sid: {
                    final BOFRecord bof = (BOFRecord) record;
                    if (BOFRecord.TYPE_WORKBOOK == bof.getType()) { // workbook starts
                        boundSheets = new HashMap<>();
                    } else if (BOFRecord.TYPE_WORKSHEET == bof.getType()) { // new sheet starts
                        handleStartSheet();
                    }
                    break;
                }
                case RowRecord.sid: {
                    final RowRecord row = (RowRecord) record;
                    assert currentSheetRows != null;

                    currentSheetRows.put(++tRowNum, row);
                    break;
                }
                case BlankRecord.sid: {
                    final BlankRecord blank = (BlankRecord) record;

                    handleCell(blank.getRow(), blank.getColumn(), null);
                    break;
                }
                case MulBlankRecord.sid: {
                    final MulBlankRecord mulBlank = (MulBlankRecord) record;

                    for (int column = mulBlank.getFirstColumn();
                         column <= mulBlank.getLastColumn();
                         column++) {
                        handleCell(mulBlank.getRow(), column, null);
                    }
                    break;
                }
                case NumberRecord.sid: {
                    final NumberRecord number = (NumberRecord) record;

                    final Object cellValue = formatNumberDateCell(number);
                    handleCell(number.getRow(), number.getColumn(), cellValue);
                    break;
                }
                case FormulaRecord.sid: {
                    final FormulaRecord formula = (FormulaRecord) record;

                    @SuppressWarnings("deprecation") final CellType resultType =
                            CellType.forInt(formula.getCachedResultType());
                    switch (resultType) {
                        case NUMERIC: {
                            final Object cellValue = formatNumberDateCell(formula);
                            handleCell(formula.getRow(), formula.getColumn(), cellValue);
                            break;
                        }
                        case STRING: {
                            assert formula.hasCachedResultString();
                            // does nothing, value will be stored in the next StringRecord
                            break;
                        }
                        case BOOLEAN: {
                            final boolean cellValue = formula.getCachedBooleanValue();
                            handleCell(formula.getRow(), formula.getColumn(), cellValue);
                            break;
                        }
                        case ERROR: {
                            handleCell(formula.getRow(), formula.getColumn(), null);
                            break;
                        }
                        default: {
                            throw new AssertionError("Shouldn't happen");
                        }
                    }
                    break;
                }
                case StringRecord.sid: {
                    final StringRecord string = (StringRecord) record;
                    if (previousRecord instanceof FormulaRecord) {
                        final FormulaRecord formula = (FormulaRecord) previousRecord;

                        final String cellValue = string.getString();
                        handleCell(formula.getRow(), formula.getColumn(), formatString(cellValue));
                    }
                    break;
                }
                case RKRecord.sid: {
                    final RKRecord rk = (RKRecord) record;

                    final double cellValue = rk.getRKNumber();
                    handleCell(rk.getRow(), rk.getColumn(), cellValue);
                    break;
                }
                case LabelRecord.sid: {
                    final LabelRecord label = (LabelRecord) record;

                    final String cellValue = label.getValue();
                    handleCell(label.getRow(), label.getColumn(), formatString(cellValue));
                    break;
                }
                case LabelSSTRecord.sid: {
                    final LabelSSTRecord labelSst = (LabelSSTRecord) record;

                    final int sstIndex = labelSst.getSSTIndex();
                    final String cellValue = sharedStringTable.getString(sstIndex)
                            .getString();
                    handleCell(labelSst.getRow(), labelSst.getColumn(), formatString(cellValue));
                    break;
                }
            }

            previousRecord = record;

            return USER_CODE_CONTINUE;
        }

        private void handleStartSheet() {
            if (!previousSheetEndHandled) {
                handleEndSheet(previousSheetIndex);
            }

            previousSheetIndex = currentSheetIndex;
            currentSheetIndex += 1;
            final BoundSheetRecord boundSheet = boundSheets.get(currentSheetIndex);
            assert boundSheet != null;

            final String sheetName = boundSheet.getSheetname();
            handler.onStartSheet(currentSheetIndex, sheetName);

            currentSheetRows = new HashMap<>();
            tRowNum = -1;
            previousRowNumber = -1;
            previousSheetEndHandled = false;
        }

        @Nullable
        private Object formatNumberDateCell(@NotNull CellValueRecordInterface cellRecord) {
            final double value;
            if (cellRecord instanceof NumberRecord) {
                value = ((NumberRecord) cellRecord).getValue();
            } else if (cellRecord instanceof FormulaRecord) {
                value = ((FormulaRecord) cellRecord).getValue();
            } else {
                throw new AssertionError("Shouldn't happen");
            }

            final int formatIndex = formatTracker.getFormatIndex(cellRecord);
            final String formatString = formatTracker.getFormatString(cellRecord);

            if (formatString != null) {
                boolean returnAsString = false;

                if (ReaderUtils.isATextFormat(formatIndex, formatString)) { // deals with cell marked as text
                    returnAsString = true;
                } else if (DateUtil.isADateFormat(formatIndex, formatString)) { // deals with date
                    if (ReaderUtils.isValidExcelDate(value)) {
                        return ReaderUtils.toJsr310DateOrTime(value, use1904Windowing);
                    } else {
                        returnAsString = true;
                    }
                }

                if (returnAsString) {
                    if (ReaderUtils.isAWholeNumber(value)) {
                        return Long.toString((long) (value));
                    } else {
                        return Double.toString(value);
                    }
                }

                final String decimalStringValue = dataFormatter.formatRawCellContents(value, formatIndex, formatString);
                return ReaderUtils.decimalStringToDecimal(decimalStringValue);
            }

            return value;
        }

        @Nullable
        private String formatString(@Nullable String value) {
            return "".equals(value) ? null : value;
        }

        private void handleEndSheet(int sheetIndex) {
            if (!previousRowEndHandled) {
                handleEndRow(currentRowNumber);
            }

            handler.onEndSheet(sheetIndex);
            previousSheetEndHandled = true;
        }

        private void handleCell(int rowNum, int columnNum, @Nullable Object cellValue) {
            previousRowNumber = currentRowNumber;
            currentRowNumber = rowNum;

            if (previousRowNumber != currentRowNumber) {
                handleStartRow();
            }

            cellValue = ReaderUtils.toRelativeType(cellValue);
            handler.onHandleCell(currentSheetIndex, rowNum, columnNum,
                    CellValue.newInstance(cellValue));

            if (currentRowEndColumnNum == columnNum) {
                handleEndRow(currentRowNumber);
            }
        }

        private void handleStartRow() {
            if (!previousRowEndHandled) {
                handleEndRow(previousRowNumber);
            }

            RowRecord row = currentSheetRows.get(currentRowNumber);
            currentRowEndColumnNum = row.getLastCol() - 1;

            handler.onStartRow(currentSheetIndex, currentRowNumber);
            previousRowEndHandled = false;
        }

        private void handleEndRow(int rowNum) {
            handler.onEndRow(currentSheetIndex, rowNum);
            previousRowEndHandled = true;
        }
    }

}
