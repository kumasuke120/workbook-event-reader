package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.WorkbookEventReader;
import com.github.kumasuke120.excel.handler.WorkbookRecord.MetadataType;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

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
     * Extracts records from the specified {@link WorkbookEventReader}.
     *
     * @param reader the reader to extract records from
     * @param <T>    the type of the record to extract
     * @return a list of extracted records
     */
    public static <T> List<T> extract(WorkbookEventReader reader, Class<T> recordClass) {
        final WorkbookRecordExtractor<T> extractor = new WorkbookRecordExtractor<>(recordClass);
        reader.read(extractor);
        return extractor.getResult();
    }

    /**
     * Returns a list of extracted records after reading using {@link WorkbookEventReader#read(WorkbookEventReader.EventHandler)}.
     *
     * @return a list of extracted records
     */
    public List<E> getResult() {
        return new ArrayList<>(result);
    }

    @Override
    public void onStartDocument() {
        result = new ArrayList<>();
    }

    @Override
    public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
        currentSheetName = sheetName;
        if (recordMapper.beyondRange(sheetIndex)) {
            WorkbookEventReader.currentRead().cancel();
        }
    }

    @Override
    public void onEndSheet(int sheetIndex) {
        currentSheetName = null;
        if (recordMapper.beyondRange(sheetIndex - 1)) {
            WorkbookEventReader.currentRead().cancel();
        }
    }

    @Override
    public void onStartRow(int sheetIndex, int rowNum) {
        if (!recordMapper.withinRange(sheetIndex, rowNum)) {
            return;
        }

        currentRecord = createRecord();

        recordMapper.setMetadata(MetadataType.SHEET_INDEX, currentRecord, sheetIndex);
        recordMapper.setMetadata(MetadataType.SHEET_NAME, currentRecord, currentSheetName);
        recordMapper.setMetadata(MetadataType.ROW_NUMBER, currentRecord, rowNum);
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
        if (!recordMapper.withinRange(sheetIndex, rowNum, columnNum)) {
            return;
        }

        recordMapper.setValue(currentRecord, columnNum, cellValue);
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
