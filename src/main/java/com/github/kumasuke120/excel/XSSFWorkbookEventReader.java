package com.github.kumasuke120.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.SAXParseException;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Iterator;
import java.util.Map;

/**
 * A {@link WorkbookEventReader} reads a SpreadsheetML workbook (Excel 2007 onwards) whose file extension
 * usually is <code>xlsx</code>
 */
@SuppressWarnings("unused")
public class XSSFWorkbookEventReader extends AbstractWorkbookEventReader {

    private static final ThreadLocal<Boolean> use1904WindowingLocal = ThreadLocal.withInitial(() -> false);

    private OPCPackage opcPackage;
    private XSSFReader xssfReader;
    private SharedStringsTable sharedStringsTable;
    private StylesTable stylesTable;

    private boolean use1904Windowing;

    /**
     * Creates a new {@link XSSFWorkbookEventReader} based on the given file path.
     *
     * @param filePath file path of the workbook
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    public XSSFWorkbookEventReader(@NotNull(exception = NullPointerException.class) Path filePath) {
        this(filePath, null);
    }

    /**
     * Creates a new {@link XSSFWorkbookEventReader} based on the given file path using given password.
     *
     * @param filePath file path of the workbook
     * @param password password to open the file
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    public XSSFWorkbookEventReader(@NotNull(exception = NullPointerException.class) Path filePath,
                                   @Nullable String password) {
        super(filePath, password);
    }

    /**
     * Creates a new {@link XSSFWorkbookEventReader} based on the given workbook {@link InputStream}
     * and the given password if possible.
     *
     * @param in {@link InputStream} of the workbook to be read
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    public XSSFWorkbookEventReader(@NotNull(exception = NullPointerException.class) InputStream in) {
        this(in, null);
    }

    /**
     * Creates a new {@link XSSFWorkbookEventReader} based on the given workbook {@link InputStream}
     * and the given password if possible.
     *
     * @param in       {@link InputStream} of the workbook to be read
     * @param password password to open the file
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    public XSSFWorkbookEventReader(@NotNull(exception = NullPointerException.class) InputStream in,
                                   @Nullable String password) {
        super(in, password);
    }

    /**
     * Sets all following-opened instances of {@link XSSFWorkbookEventReader} on the current thread using
     * 1904 windowing for parsing date cells whether or not.<br>
     *
     * @param use1904Windowing whether use 1904 date windowing or not
     */
    public static void setUse1904Windowing(boolean use1904Windowing) {
        use1904WindowingLocal.set(use1904Windowing);
    }

    @Override
    void doOnStartOpen() {
        use1904Windowing = use1904WindowingLocal.get();
    }

    @Override
    void doOpen(@NotNull InputStream in, @Nullable String password) throws Exception {
        Exception thrown = null;

        InputStream stream = null;
        try {
            if (password == null) {
                stream = in;
            } else {
                final POIFSFileSystem fs = new POIFSFileSystem(in);
                stream = DocumentFactoryHelper.getDecryptedStream(fs, password);
            }

            opcPackage = OPCPackage.open(stream); // consumes all stream data to memory
            initFromOpcPackage();
        } catch (Exception e) {
            thrown = e;
        } finally {
            suppressClose(stream, thrown);
        }
    }

    @Override
    void doOpen(@NotNull Path filePath, @Nullable String password) throws Exception {
        final File file = filePath.toFile();

        final InputStream stream;
        if (password != null) {
            try (final POIFSFileSystem fs = new POIFSFileSystem(file, true)) {
                stream = DocumentFactoryHelper.getDecryptedStream(fs, password);
                doOpen(stream, null);
            }
        } else {
            opcPackage = OPCPackage.open(file, PackageAccess.READ);
            initFromOpcPackage();
        }
    }

    private void initFromOpcPackage() throws IOException, OpenXML4JException {
        xssfReader = new XSSFReader(opcPackage);
        sharedStringsTable = xssfReader.getSharedStringsTable();
        stylesTable = xssfReader.getStylesTable();
    }

    @Override
    @NotNull
    ReaderCleanAction createCleanAction() {
        return new XSSFReaderCleanAction(this);
    }

    @Override
    void doRead(@NotNull EventHandler handler) throws Exception {
        int currentSheetIndex = -1;

        handler.onStartDocument();

        final SAXParser saxParser = createSAXParser();
        final ReaderSheetHandler saxHandler = new ReaderSheetHandler(handler);

        final XSSFReader.SheetIterator sheetIt = getSheetIterator();
        while (sheetIt.hasNext()) {
            if (!isReading()) {
                handler.onReadCancelled(); // has next sheet but reading is false
                break;
            }

            try (final InputStream sheetIs = sheetIt.next()) {
                String sheetName = sheetIt.getSheetName();
                handler.onStartSheet(++currentSheetIndex, sheetName);

                saxHandler.initializeForNewSheet(currentSheetIndex);
                try {
                    saxParser.parse(sheetIs, saxHandler);
                } catch (CancelReadingException e) {
                    handler.onReadCancelled();
                    // stops parsing and cancels reading
                    break;
                }

                handler.onEndSheet(currentSheetIndex);
            }
        }

        if (isReading()) { // only triggers the event when the reading process wasn't cancelled
            handler.onEndDocument();
        }
    }

    @NotNull
    private XSSFReader.SheetIterator getSheetIterator() throws IOException, InvalidFormatException {
        final Iterator<InputStream> sheetsData = xssfReader.getSheetsData();
        assert sheetsData instanceof XSSFReader.SheetIterator;
        return (XSSFReader.SheetIterator) sheetsData;
    }

    @NotNull
    private SAXParser createSAXParser() throws ParserConfigurationException, SAXException {
        final SAXParserFactory factory = SAXParserFactory.newInstance();
        factory.setNamespaceAware(true);
        return factory.newSAXParser();
    }

    // stops EventHandler, not an actual exception
    private static class CancelReadingException extends SAXException {
    }

    private static class XSSFReaderCleanAction extends ReaderCleanAction {
        private final OPCPackage opcPackage;

        XSSFReaderCleanAction(@NotNull XSSFWorkbookEventReader reader) {
            this.opcPackage = reader.opcPackage;
        }

        @Override
        void doClean() throws Exception {
            if (opcPackage != null) {
                opcPackage.close();
            }
        }
    }

    private class ReaderSheetHandler extends DefaultHandler {
        // tags
        private static final String TAG_ROW = "row";
        private static final String TAG_CELL = "c";
        private static final String TAG_CELL_VALUE = "v";
        private static final String TAG_INLINE_STR = "is";
        private static final String TAG_INLINE_CELL_VALUE = "t";

        // attribute for TAG_ROW
        private static final String ATTRIBUTE_ROW_REFERENCE = "r";

        // attributes for TAG_CELL
        private static final String ATTRIBUTE_CELL_REFERENCE = "r";
        private static final String ATTRIBUTE_CELL_TYPE = "t";
        private static final String ATTRIBUTE_CELL_STYLE = "s";

        // attribute values for ATTRIBUTE_CELL_REFERENCE
        private static final String CELL_TYPE_SHARED_STRING = "s";
        private static final String CELL_TYPE_INLINE_STRING = "inlineStr";
        private static final String CELL_TYPE_STRING = "str";
        private static final String CELL_TYPE_BOOLEAN = "b";
        private static final String CELL_TYPE_ERROR = "e";

        // cell values for boolean
        private static final String CELL_VALUE_BOOLEAN_TRUE = "1";
        private static final String CELL_VALUE_BOOLEAN_FALSE = "0";

        private final EventHandler handler;
        private final StringBuilder currentCellValueBuilder = new StringBuilder();

        private String currentElementQName;

        private int currentSheetIndex = -1;
        private int currentRowNum = -1;
        private int currentColumnNum = -1;

        private int currentCellXfIndex = -1;
        private String currentCellType;

        private boolean isCurrentCellValue = false;

        ReaderSheetHandler(@NotNull EventHandler handler) {
            this.handler = handler;
        }

        void initializeForNewSheet(int currentSheetIndex) {
            this.currentSheetIndex = currentSheetIndex;
            this.currentRowNum = -1;
        }

        @Override
        public final void startElement(@NotNull String uri, @NotNull String localName,
                                       @NotNull String qName, @NotNull Attributes attributes) throws SAXException {
            cancelReadingWhenNecessary();

            currentElementQName = qName;

            if (TAG_CELL.equals(localName)) {
                final String currentCellReference = attributes.getValue(ATTRIBUTE_CELL_REFERENCE);
                if (currentCellReference == null) { // cell reference can be left out
                    // treats as a continuation to previous column
                    if (currentColumnNum == -1) {
                        currentColumnNum = 0;
                    } else {
                        currentColumnNum += 1;
                    }
                } else {
                    final Map.Entry<Integer, Integer> rowAndColumn =
                            Util.cellReferenceToRowAndColumn(currentCellReference);

                    if (rowAndColumn == null) {
                        throw new SAXParseException(
                                "Cannot parse row number or column number in tag '" + qName + "'", null);
                    }

                    final int rowNum = rowAndColumn.getKey();
                    final int columnNum = rowAndColumn.getValue();

                    assert rowNum == currentRowNum;
                    currentColumnNum = columnNum;
                }

                // saves styles of current cell
                currentCellXfIndex = Util.toInt(attributes.getValue(ATTRIBUTE_CELL_STYLE), -1);
                currentCellType = attributes.getValue(ATTRIBUTE_CELL_TYPE);
            } else if (TAG_ROW.equals(localName)) {
                final String rawValue = attributes.getValue(ATTRIBUTE_ROW_REFERENCE);
                if (rawValue == null) { // row reference can be left out
                    // treats as a continuation to previous row
                    if (currentRowNum == -1) {
                        currentRowNum = 0;
                    } else {
                        currentRowNum += 1;
                    }
                } else {
                    try {
                        currentRowNum = Integer.parseInt(attributes.getValue(ATTRIBUTE_ROW_REFERENCE)) - 1;
                    } catch (NumberFormatException e) {
                        throw new SAXParseException("Cannot parse row number in tag '" + qName + "'",
                                                    null, e);
                    }
                }

                handler.onStartRow(currentSheetIndex, currentRowNum);
            } else if (TAG_INLINE_STR.equals(localName)) {
                if (currentCellType == null) {
                    currentCellType = CELL_TYPE_INLINE_STRING;
                }
            } else if (isCellValueRelated(localName)) {
                isCurrentCellValue = true; // indicates cell value starts
                currentCellValueBuilder.setLength(0);
            }
        }

        @Override
        public final void endElement(@NotNull String uri, @NotNull String localName,
                                     @NotNull String qName) throws SAXException {
            cancelReadingWhenNecessary();

            currentElementQName = qName;

            if (TAG_CELL.equals(localName)) {
                final Object cellValue = getCurrentCellValue();
                handler.onHandleCell(currentSheetIndex, currentRowNum, currentColumnNum,
                                     CellValue.newInstance(cellValue));

                // clear its content after processing
                currentCellValueBuilder.setLength(0);
                currentCellType = null;
            } else if (TAG_ROW.equals(localName)) {
                currentColumnNum = -1;
                handler.onEndRow(currentSheetIndex, currentRowNum);
            } else if (isCellValueRelated(localName)) {
                isCurrentCellValue = false; // indicates cell value ends
            }
        }

        private boolean isCellValueRelated(@NotNull String localName) {
            return TAG_CELL_VALUE.equals(localName) ||
                    (CELL_TYPE_INLINE_STRING.equals(currentCellType) && TAG_INLINE_CELL_VALUE.equals(localName));
        }

        @Nullable
        private Object getCurrentCellValue() throws SAXParseException {
            final Object cellValue;
            if (CELL_TYPE_ERROR.equals(currentCellType)) {
                cellValue = null;
            } else if (CELL_TYPE_BOOLEAN.equals(currentCellType)) {
                cellValue = getCurrentBooleanCellValue();
            } else { // treats all other cases as string firstly
                final String stringCellValue = getCurrentStringCellValue();
                cellValue = formatNumberDateCellValue(stringCellValue);
            }

            return Util.toRelativeType(cellValue);
        }

        @NotNull
        private Boolean getCurrentBooleanCellValue() throws SAXParseException {
            final String currentCellValue = currentCellValueBuilder.toString();
            if (CELL_VALUE_BOOLEAN_TRUE.equals(currentCellValue)) {
                return Boolean.TRUE;
            } else if (CELL_VALUE_BOOLEAN_FALSE.equals(currentCellValue)) {
                return Boolean.FALSE;
            } else {
                throw new SAXParseException("Cannot parse boolean value in tag '" + currentElementQName + "', " +
                                                    "which should be 'TRUE' or 'FALSE': " + currentCellValue,
                                            null);
            }
        }

        @Nullable
        private String getCurrentStringCellValue() throws SAXParseException {
            if (CELL_TYPE_SHARED_STRING.equals(currentCellType)) {
                return getCurrentSharedStringCellValue();
            } else {
                return currentCellValueBuilder.toString();
            }
        }

        @Nullable
        private String getCurrentSharedStringCellValue() throws SAXParseException {
            final String currentCellValue = currentCellValueBuilder.toString();

            final int sharedStringIndex;
            try {
                sharedStringIndex = Integer.parseInt(currentCellValue);
            } catch (NumberFormatException e) {
                throw new SAXParseException(
                        "Cannot parse shared string index in tag '" + currentElementQName + "', " +
                                "which should be a int: " + currentCellValue,
                        null, e);
            }

            final RichTextString sharedString = sharedStringsTable.getItemAt(sharedStringIndex);
            return sharedString.getString();
        }

        @Nullable
        private Object formatNumberDateCellValue(@Nullable String stringCellValue) {
            final Object cellValue;

            // valid numFmtId is non-negative, -1 denotes there is no cell format for the cell
            final short formatIndex = currentCellXfIndex == -1 ? -1 : getFormatIndex(currentCellXfIndex);
            final String formatString = currentCellXfIndex == -1 ? null : getFormatString(formatIndex);

            if (stringCellValue == null || stringCellValue.isEmpty()) {
                cellValue = null;
            } else if (isCurrentCellString() ||
                    Util.isATextFormat(formatIndex, formatString)) { // deals with cell marked as text
                cellValue = stringCellValue;
            } else if (DateUtil.isADateFormat(formatIndex, formatString)) { // deals with date format
                Object theValue;
                try {
                    double doubleValue = Double.parseDouble(stringCellValue);
                    if (Util.isValidExcelDate(doubleValue)) {
                        theValue = Util.toJsr310DateOrTime(doubleValue, use1904Windowing);
                    } else {
                        // treats invalid value as text
                        theValue = stringCellValue;
                    }
                } catch (NumberFormatException e) {
                    // non-double value in a date format cell, which is tolerable
                    theValue = stringCellValue;
                }
                cellValue = theValue;
            } else if (Util.isAWholeNumber(stringCellValue)) { // deals with whole number
                // will never throw NumberFormatException
                cellValue = Long.parseLong(stringCellValue);
            } else if (Util.isADecimalFraction(stringCellValue)) { // deals with decimal fraction
                // will never throw NumberFormatException
                cellValue = Double.parseDouble(stringCellValue);
            } else {
                cellValue = stringCellValue;
            }

            return cellValue;
        }

        private short getFormatIndex(int cellXfIndex) {
            CTXf cellXf = stylesTable.getCellXfAt(cellXfIndex);
            if (cellXf.isSetNumFmtId()) {
                return (short) cellXf.getNumFmtId();
            } else {
                // valid numFmtId is non-negative, -1 denotes invalid value
                return -1;
            }
        }

        @NotNull
        private String getFormatString(short numFmtId) {
            String numberFormat = stylesTable.getNumberFormatAt(numFmtId);
            if (numberFormat == null) {
                numberFormat = BuiltinFormats.getBuiltinFormat(numFmtId);
            }
            return numberFormat;
        }

        private boolean isCurrentCellString() {
            return CELL_TYPE_INLINE_STRING.equals(currentCellType) ||
                    CELL_TYPE_SHARED_STRING.equals(currentCellType) ||
                    CELL_TYPE_STRING.equals(currentCellType);
        }

        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            cancelReadingWhenNecessary();

            if (isCurrentCellValue) { // only records when cell value starts
                currentCellValueBuilder.append(ch, start, length);
            }
        }

        private void cancelReadingWhenNecessary() throws SAXException {
            if (!isReading()) {
                throw new CancelReadingException();
            }
        }
    }

}
