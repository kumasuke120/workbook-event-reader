package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import com.github.kumasuke120.excel.util.CollectionUtils;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
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

    private final Class<E> recordClass;

    private final WorkbookRecordMapper<E> recordMapper;

    private final Constructor<E> recordConstructor;

    private Map<Integer, String> columnTitles;

    private List<E> result;

    private String currentSheetName;

    private E currentRecord;

    /**
     * Constructs a new {@code WorkbookRecordExtractor} for the specified record class.
     *
     * @param recordClass the record class
     * @throws WorkbookRecordException if the specified class is not annotated with {@code @WorkbookRecord}
     */
    public WorkbookRecordExtractor(@NotNull(exception = NullPointerException.class) Class<E> recordClass) {
        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.recordMapper = new WorkbookRecordMapper<>(recordClass);
        this.recordConstructor = findNoArgRecordConstructor();
    }

    private Constructor<E> findNoArgRecordConstructor() {
        try {
            return recordClass.getConstructor();
        } catch (NoSuchMethodException e) {
            return null;
        }
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
        return getResult();
    }

    /**
     * Returns a list of extracted records after reading using {@link WorkbookEventReader#read(WorkbookEventReader.EventHandler)}.
     *
     * @return a list of extracted records
     */
    public List<E> getResult() {
        return new ArrayList<>(result);
    }

    /**
     * Returns the title of the specified column number.
     *
     * @param columnNum the column number
     * @return the title of the specified column number
     */
    public String getColumnTitle(int columnNum) {
        if (CollectionUtils.isEmpty(columnTitles)) {
            return null;
        }
        final String title = columnTitles.get(columnNum);
        return title == null ? "" : title;
    }

    /**
     * Returns a list of all column titles.
     *
     * @return a list of all column titles
     */
    public List<String> getColumnTitles() {
        if (CollectionUtils.isEmpty(columnTitles)) {
            return null;
        }
        final Integer maxColumn = Collections.max(columnTitles.keySet());
        if (maxColumn == null) {
            return null;
        }
        final List<String> titles = new ArrayList<>(maxColumn + 1);
        for (int i = 0; i <= maxColumn; i++) {
            titles.add(getColumnTitle(i));
        }
        return titles;
    }

    @Override
    public void onStartDocument() {
        columnTitles = new TreeMap<>();
        result = new ArrayList<>();
    }

    @Override
    public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
        currentSheetName = sheetName;
        cancelIfSheetBeyondRange(sheetIndex);
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
            result.add(currentRecord);
        }
    }

    @Override
    public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
        if (recordMapper.isTitleRow(rowNum)) {
            if (!HandlerUtils.isValueEmpty(cellValue)) {
                String columnTitle = cellValue.trim().stringValue();
                columnTitles.put(columnNum, columnTitle);
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
        if (recordConstructor != null) {
            try {
                return recordConstructor.newInstance();
            } catch (InstantiationException | IllegalAccessException e) {
                throw new WorkbookRecordException("cannot initialize @WorkbookRecord record class", e);
            } catch (InvocationTargetException e) {
                throw new WorkbookRecordException("error encountered when initializing @WorkbookRecord record class",
                        e.getTargetException());
            }
        }

        Exception e = new UnsupportedOperationException("or you can override WorkbookRecordExtractor#createRecord()");
        throw new WorkbookRecordException("@WorkbookRecord record class has no suitable constructor", e);
    }

}
