package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.util.CollectionUtils;
import com.github.kumasuke120.excel.util.ObjectCreationException;
import com.github.kumasuke120.excel.util.ObjectFactory;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.util.*;

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

    private Map<Integer, List<E>> sheetRecords;

    private String currentSheetName;

    private E currentRecord;

    /**
     * Constructs a new {@code WorkbookRecordExtractor} for the specified record class.
     *
     * @param recordClass the record class
     * @throws WorkbookRecordException if the specified class is not annotated with {@code @WorkbookRecord}
     */
    public WorkbookRecordExtractor(@NotNull(exception = NullPointerException.class) Class<E> recordClass) {
        HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.recordMapper = new WorkbookRecordMapper<>(recordClass);
        this.recordFactory = ObjectFactory.buildFactory(recordClass);
    }

    /**
     * Creates a new {@code WorkbookRecordExtractor} for the specified record class.
     *
     * @param recordClass the record class
     * @param <T>         the type of the record to extract
     * @return a new {@code WorkbookRecordExtractor} for the specified record class
     */
    public static <T> WorkbookRecordExtractor<T> ofRecord(Class<T> recordClass) {
        return new WorkbookRecordExtractor<>(recordClass);
    }

    /**
     * Extracts records from the specified {@link WorkbookEventReader}.
     *
     * @param reader the {@link WorkbookEventReader} to read
     * @return a list of extracted records
     */
    public List<E> extract(WorkbookEventReader reader) {
        reader.read(this);
        return getAllRecords();
    }

    /**
     * Returns a list of all records in all sheets.
     *
     * @return a list of all records
     */
    public List<E> getAllRecords() {
        if (CollectionUtils.isEmpty(sheetRecords)) {
            return null;
        }

        final List<E> records = new ArrayList<>();
        for (List<E> sheetRecord : sheetRecords.values()) {
            records.addAll(sheetRecord);
        }
        return records;
    }

    /**
     * Returns a list of records in the specified sheet.
     *
     * @param sheetIndex the index of the sheet
     * @return a list of records in the specified sheet
     */
    public List<E> getRecords(int sheetIndex) {
        if (CollectionUtils.isEmpty(sheetRecords)) {
            return null;
        }
        final List<E> records = sheetRecords.get(sheetIndex);
        if (records == null) {
            return null;
        }
        return new ArrayList<>(records);
    }

    /**
     * Returns the title of the specified column in given sheet.
     *
     * @param sheetIndex the index of the sheet
     * @param columnNum  the column number
     * @return the title of the specified column number
     */
    public String getColumnTitle(int sheetIndex, int columnNum) {
        Map<Integer, String> sheetColumnTitles;
        if (CollectionUtils.isEmpty(columnTitles) ||
                (sheetColumnTitles = columnTitles.get(sheetIndex)) == null) {
            return null;
        }
        final String title = sheetColumnTitles.get(columnNum);
        return title == null ? "" : title;
    }

    /**
     * Returns a list of all column titles in given sheet.
     *
     * @param sheetIndex the index of the sheet
     * @return a list of all column titles
     */
    public List<String> getAllColumnTitles(int sheetIndex) {
        if (CollectionUtils.isEmpty(columnTitles)) {
            return null;
        }
        Map<Integer, String> sheetColumnTitles = columnTitles.get(sheetIndex);
        if (sheetColumnTitles == null) {
            return null;
        }
        final Integer maxColumn;
        try {
            maxColumn = Collections.max(sheetColumnTitles.keySet());
        } catch (NoSuchElementException e) {
            return null;
        }
        final List<String> titles = new ArrayList<>(maxColumn + 1);
        for (int i = 0; i <= maxColumn; i++) {
            titles.add(getColumnTitle(sheetIndex, i));
        }
        return titles;
    }

    @Override
    public void onStartDocument() {
        columnTitles = new TreeMap<>();
        sheetRecords = new TreeMap<>();
    }


    @Override
    public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
        currentSheetName = sheetName;
        cancelIfSheetBeyondRange(sheetIndex);

        columnTitles.putIfAbsent(sheetIndex, recordMapper.getDefaultColumnTitles());
        sheetRecords.putIfAbsent(sheetIndex, new ArrayList<>());
    }

    @Override
    public void onEndSheet(int sheetIndex) {
        currentSheetName = null;
        cancelIfSheetBeyondRange(sheetIndex + 1);
    }

    void cancelIfSheetBeyondRange(int sheetIndex) {
        if (recordMapper.beyondRange(sheetIndex)) {
            WorkbookEventReader.currentRead().cancel();
        }
    }

    @Override
    public void onStartRow(int sheetIndex, int rowNum) {
        if (!recordMapper.withinRange(sheetIndex, rowNum)) {
            return;
        }

        currentRecord = createRecord();

        recordMapper.setMetadata(currentRecord, MetadataType.SHEET_INDEX, sheetIndex);
        recordMapper.setMetadata(currentRecord, MetadataType.SHEET_NAME, currentSheetName);
        recordMapper.setMetadata(currentRecord, MetadataType.ROW_NUMBER, rowNum);
    }

    @Override
    public void onEndRow(int sheetIndex, int rowNum) {
        if (currentRecord == null) {
            return;
        }

        if (recordMapper.withinRange(sheetIndex, rowNum)) {
            sheetRecords.get(sheetIndex).add(currentRecord);
        }
    }

    @Override
    public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
        if (recordMapper.isTitleRow(rowNum)) {
            if (!HandlerUtils.isValueEmpty(cellValue)) {
                String columnTitle = cellValue.trim().stringValue();
                columnTitles.get(sheetIndex).putIfAbsent(columnNum, columnTitle);
            }
        }

        if (recordMapper.withinRange(sheetIndex, rowNum, columnNum)) {
            recordMapper.setValue(currentRecord, columnNum, cellValue);
        }
    }

    /**
     * Creates a new record instance.
     *
     * @return a new record instance
     */
    protected E createRecord() {
        if (recordFactory != null) {
            try {
                return recordFactory.newInstance();
            } catch (ObjectCreationException e) {
                throw new WorkbookRecordException("cannot initialize @WorkbookRecord record class", e);
            }
        }

        Exception e = new UnsupportedOperationException("or you can override WorkbookRecordExtractor#createRecord()");
        throw new WorkbookRecordException("@WorkbookRecord record class has no suitable constructor", e);
    }

}
