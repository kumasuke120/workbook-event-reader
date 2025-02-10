package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.util.CollectionUtils;
import com.github.kumasuke120.excel.util.ObjectCreationException;
import com.github.kumasuke120.excel.util.ObjectFactory;
import com.github.kumasuke120.excel.util.StringUtils;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.*;
import java.util.stream.Collectors;

/**
 * An extractor for extracting records from a workbook. It must be used with {@link WorkbookEventReader}.
 * <p>
 * This class is currently experimental and may be changed or removed in the future.
 * </p>
 *
 * @param <E> the type of the record to extract
 */
@ApiStatus.Experimental
public class WorkbookRecordExtractor<E> implements WorkbookEventReader.EventHandler {

    private final WorkbookRecordMapper<E> recordMapper;

    private final ObjectFactory<E> recordFactory;

    private Map<Integer, Map<Integer, String>> columnTitles;

    private Map<Integer, List<ExtractResult<E>>> sheetResults;

    private Map<Integer, List<ExtractResult<E>>> sheetFailResults;

    private String currentSheetName;

    private ExtractResult<E> currentRowResult;

    /**
     * Constructs a new {@code WorkbookRecordExtractor} for the specified record class.
     *
     * @param recordClass the record class
     * @throws WorkbookRecordException if the specified class is not annotated with {@code @WorkbookRecord}
     * @throws WorkbookRecordException if the specified class has no suitable no-arg constructor
     */
    public WorkbookRecordExtractor(@NotNull(exception = NullPointerException.class) Class<E> recordClass) {
        HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.recordMapper = new WorkbookRecordMapper<>(recordClass);
        try {
            this.recordFactory = ObjectFactory.buildFactory(recordClass);
        } catch (ObjectCreationException e) {
            String msg = "the record class '" + recordClass.getName() + "' lacks a suitable no-argument constructor";
            throw new WorkbookRecordException(msg, e);
        }
    }

    /**
     * Creates a new {@code WorkbookRecordExtractor} for the specified record class.
     *
     * @param recordClass the record class
     * @param <T>         the type of the record to extract
     * @return a new {@code WorkbookRecordExtractor} for the specified record class
     */
    @NotNull
    public static <T> WorkbookRecordExtractor<T> ofRecord(@NotNull Class<T> recordClass) {
        return new WorkbookRecordExtractor<>(recordClass);
    }

    /**
     * Extracts records from the specified {@link WorkbookEventReader}.
     *
     * @param reader the {@link WorkbookEventReader} to read
     * @return a list of extracted records
     */
    @NotNull
    public final List<E> extract(WorkbookEventReader reader) {
        reader.read(this);
        return getAllRecords();
    }

    /**
     * Gets all failed extract results.
     *
     * @return a list of all failed extract results
     */
    @NotNull
    public final List<ExtractResult<E>> getAllFailExtractResults() {
        if (CollectionUtils.isEmpty(sheetFailResults)) {
            return new ArrayList<>(0);
        }

        return sheetFailResults.values().stream()
                .flatMap(Collection::stream)
                .collect(Collectors.toList());
    }

    /**
     * Gets all failed extract results in the specified sheet.
     *
     * @param sheetIndex the index of the sheet
     * @return a list of all failed extract results in the specified sheet
     */
    @NotNull
    public final List<ExtractResult<E>> getFailExtractResults(int sheetIndex) {
        return getExtractResults(sheetFailResults, sheetIndex);
    }

    @Nullable
    public final ExtractResult<E> getFailExtractResult(int sheetIndex, int rowNum) {
        return getFailExtractResults(sheetIndex).stream()
                .filter(r -> rowNum == r.getRowNum())
                .findFirst()
                .orElse(null);
    }

    /**
     * Returns a list of all success records in all sheets.
     *
     * @return a list of all success records
     */
    @NotNull
    public final List<E> getAllRecords() {
        if (CollectionUtils.isEmpty(sheetResults)) {
            return new ArrayList<>(0);
        }

        return sheetResults.values().stream()
                .flatMap(Collection::stream)
                .map(ExtractResult::get)
                .collect(Collectors.toList());
    }

    /**
     * Returns a list of success records in the specified sheet.
     *
     * @param sheetIndex the index of the sheet
     * @return a list of success records in the specified sheet
     */
    @NotNull
    public final List<E> getRecords(int sheetIndex) {
        return getExtractResults(sheetResults, sheetIndex).stream()
                .map(ExtractResult::get)
                .collect(Collectors.toList());
    }

    private List<ExtractResult<E>> getExtractResults(Map<Integer, List<ExtractResult<E>>> sheetResults, int sheetIndex) {
        if (CollectionUtils.isEmpty(sheetResults)) {
            return new ArrayList<>(0);
        }
        final List<ExtractResult<E>> extractResults = sheetResults.get(sheetIndex);
        if (extractResults == null) {
            return new ArrayList<>(0);
        }
        return extractResults;
    }

    /**
     * Returns the title of the specified column in given sheet.
     *
     * @param sheetIndex the index of the sheet
     * @param columnNum  the column number
     * @return the title of the specified column number
     */
    @NotNull
    public final String getColumnTitle(int sheetIndex, int columnNum) {
        Map<Integer, String> sheetColumnTitles;
        if (CollectionUtils.isEmpty(columnTitles) ||
                (sheetColumnTitles = columnTitles.get(sheetIndex)) == null) {
            return "";
        }
        return sheetColumnTitles.getOrDefault(columnNum, "");
    }

    /**
     * Returns a list of all column titles in given sheet.
     *
     * @param sheetIndex the index of the sheet
     * @return a list of all column titles
     */
    @NotNull
    public final List<String> getAllColumnTitles(int sheetIndex) {
        if (CollectionUtils.isEmpty(columnTitles)) {
            return new ArrayList<>(0);
        }
        Map<Integer, String> sheetColumnTitles = columnTitles.get(sheetIndex);
        if (sheetColumnTitles == null) {
            return new ArrayList<>(0);
        }
        final Integer maxColumn;
        try {
            maxColumn = Collections.max(sheetColumnTitles.keySet());
        } catch (NoSuchElementException e) {
            return new ArrayList<>(0);
        }
        final List<String> titles = new ArrayList<>(maxColumn + 1);
        for (int i = 0; i <= maxColumn; i++) {
            titles.add(getColumnTitle(sheetIndex, i));
        }
        return titles;
    }

    @Override
    public final void onStartDocument() {
        columnTitles = new TreeMap<>();
        sheetResults = new TreeMap<>();
        sheetFailResults = new TreeMap<>();
    }

    @Override
    public final void onStartSheet(int sheetIndex, @NotNull String sheetName) {
        currentSheetName = sheetName;
        cancelIfSheetBeyondRange(sheetIndex);

        if (recordMapper.withinRange(sheetIndex)) {
            columnTitles.putIfAbsent(sheetIndex, recordMapper.getPresetColumnTitles());
            sheetResults.putIfAbsent(sheetIndex, new ArrayList<>());
            sheetFailResults.putIfAbsent(sheetIndex, new ArrayList<>());
        }
    }

    @Override
    public final void onEndSheet(int sheetIndex) {
        currentSheetName = null;
        cancelIfSheetBeyondRange(sheetIndex + 1);
    }

    private void cancelIfSheetBeyondRange(int sheetIndex) {
        if (recordMapper.beyondRange(sheetIndex)) {
            WorkbookEventReader.currentRead().cancel();
        }
    }

    @Override
    public final void onStartRow(int sheetIndex, int rowNum) {
        if (!recordMapper.withinRange(sheetIndex, rowNum)) {
            return;
        }

        final E currentRecord = createRecord();
        currentRowResult = new ExtractResult<>(rowNum, currentRecord);

        recordMapper.setMetadata(currentRecord, MetadataType.SHEET_INDEX, sheetIndex);
        recordMapper.setMetadata(currentRecord, MetadataType.SHEET_NAME, currentSheetName);
        recordMapper.setMetadata(currentRecord, MetadataType.ROW_NUMBER, rowNum);
    }

    @Override
    public final void onEndRow(int sheetIndex, int rowNum) {
        if (currentRowResult == null) {
            return;
        }

        if (recordMapper.withinRange(sheetIndex, rowNum)) {
            // direct failure result
            if (failCurrentIfIncomplete(sheetIndex)) {
                return;
            }
            // user-defined failure result
            currentRowResult.userFailReason = validate(currentRowResult);
            if (failCurrentIfIncomplete(sheetIndex)) {
                return;
            }

            // success result
            sheetResults.get(sheetIndex).add(currentRowResult);
        }
    }

    private boolean failCurrentIfIncomplete(int sheetIndex) {
        if (currentRowResult.incomplete()) {
            sheetFailResults.get(sheetIndex).add(currentRowResult);
            return true;
        }
        return false;
    }

    @Override
    public final void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
        if (recordMapper.withinRange(sheetIndex) && recordMapper.isTitleRow(rowNum)) {
            if (!HandlerUtils.isValueEmpty(cellValue)) {
                String columnTitle = cellValue.trim().stringValue();
                columnTitles.get(sheetIndex).putIfAbsent(columnNum, columnTitle);
            }
        }

        if (recordMapper.withinRange(sheetIndex, rowNum, columnNum)) {
            if (!cellValue.isNull()) {
                currentRowResult.setNonBlankRow();
            }

            final E currentRecord = currentRowResult.get();
            try {
                recordMapper.setValue(currentRecord, columnNum, cellValue);
            } catch (WorkbookRecordException e) {
                recordSetValueError(sheetIndex, columnNum, cellValue, e);
            }
        }
    }

    private void recordSetValueError(int sheetIndex, int columnNum, @NotNull CellValue cellValue,
                                     @NotNull WorkbookRecordException e) {
        final String columnTitle = getColumnTitle(sheetIndex, columnNum);
        final String propertyName = recordMapper.getPropertyName(columnNum);
        ExtractColumnError columnError = new ExtractColumnError(columnNum, propertyName, columnTitle, cellValue, e);
        currentRowResult.addColumnError(columnError);
    }

    /**
     * Creates a new instance of a record.
     * <ul>
     * <li>This method is invoked when a new record needs to be created.</li>
     * <li>The default implementation uses the no-arg constructor of the record class to create a new instance.</li>
     * <li>You can override this method to implement a custom creation process and even perform additional initializations on the created record.</li>
     * </ul>
     *
     * @return a newly created record instance
     */
    protected E createRecord() {
        try {
            return recordFactory.newInstance();
        } catch (ObjectCreationException e) {
            throw new WorkbookRecordException("cannot initialize @WorkbookRecord record class", e);
        }
    }

    /**
     * Validates the current extract result.
     *
     * @param extractResult the extract result, it must wrap a fully constructed record
     * @return the user fail reason, if it is not empty, the extract result will be treated as a failure
     */
    @NotNull
    protected String validate(@NotNull ExtractResult<E> extractResult) {
        return "";
    }

    /**
     * Represents an extract result.
     *
     * @param <T> the type of the record to extract
     */
    public static class ExtractResult<T> {

        private final int rowNum;
        private final T record;
        private boolean blankRow;
        private String userFailReason;
        private final List<ExtractColumnError> columnErrors;

        ExtractResult(int rowNum, T record) {
            this.rowNum = rowNum;
            this.record = record;
            this.blankRow = true;
            this.userFailReason = "";
            this.columnErrors = new ArrayList<>();
        }

        public int getRowNum() {
            return rowNum;
        }

        public T get() {
            return record;
        }

        public List<ExtractColumnError> getColumnErrors() {
            return new ArrayList<>(columnErrors);
        }

        public String getUserFailReason() {
            return userFailReason;
        }

        void setNonBlankRow() {
            this.blankRow = false;
        }

        void addColumnError(ExtractColumnError columnError) {
            columnErrors.add(columnError);
        }

        boolean incomplete() {
            return blankRow || CollectionUtils.isNotEmpty(columnErrors) || StringUtils.isNotEmpty(userFailReason);
        }
    }

    public enum ExtractErrorKind {

        BLANK_ROW,

        COLUMN_ERRORS,

        USER_VALIDATION_FAIL

    }

    public static class ExtractColumnError {

        private final int columnNum;
        private final String propertyName;
        private final String title;
        private final CellValue cellValue;
        private final WorkbookRecordException exception;

        public ExtractColumnError(int columnNum, String propertyName,
                                  String title, CellValue cellValue,
                                  WorkbookRecordException exception) {
            this.columnNum = columnNum;
            this.propertyName = propertyName;
            this.title = title;
            this.cellValue = cellValue;
            this.exception = exception;
        }

        public int getColumnNum() {
            return columnNum;
        }

        public String getPropertyName() {
            return propertyName;
        }

        public String getTitle() {
            return title;
        }

        public CellValue getCellValue() {
            return cellValue;
        }

        public Exception getException() {
            return exception;
        }
    }


}
