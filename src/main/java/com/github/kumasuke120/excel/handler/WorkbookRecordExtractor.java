package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.WorkbookEventReader;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

@ApiStatus.Experimental
public class WorkbookRecordExtractor<E> implements WorkbookEventReader.EventHandler {

    private final Class<E> recordClass;

    private final WorkbookRecordBinder<E> recordBinder;

    private final Constructor<E> constructor;

    private List<E> result;

    private E currentRecord;

    public WorkbookRecordExtractor(@NotNull(exception = NullPointerException.class) Class<E> recordClass) {

        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.recordBinder = new WorkbookRecordBinder<>(recordClass);
        this.constructor = findNoArgConstructor();
    }

    private Constructor<E> findNoArgConstructor() {
        try {
            return recordClass.getConstructor();
        } catch (NoSuchMethodException e) {
            return null;
        }
    }

    public static <T> List<T> extract(WorkbookEventReader reader, Class<T> recordClass) {
        final WorkbookRecordExtractor<T> extractor = new WorkbookRecordExtractor<>(recordClass);
        reader.read(extractor);
        return extractor.getResult();
    }

    public List<E> getResult() {
        return new ArrayList<>(result);
    }

    @Override
    public void onStartDocument() {
        result = new ArrayList<>();
    }

    @Override
    public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
        if (recordBinder.beyondRange(sheetIndex)) {
            WorkbookEventReader.currentRead().cancel();
        }
    }

    @Override
    public void onEndSheet(int sheetIndex) {
        if (recordBinder.beyondRange(sheetIndex - 1)) {
            WorkbookEventReader.currentRead().cancel();
        }
    }

    @Override
    public void onStartRow(int sheetIndex, int rowNum) {
        if (!recordBinder.withinRange(sheetIndex, rowNum)) {
            return;
        }

        currentRecord = createRecord();

        recordBinder.setSheetIndex(currentRecord, sheetIndex);
        recordBinder.setRowNumber(currentRecord, rowNum);
    }

    @Override
    public void onEndRow(int sheetIndex, int rowNum) {
        if (currentRecord == null) {
            return;
        }

        if (recordBinder.withinRange(sheetIndex, rowNum)) {
            result.add(currentRecord);
        }
    }

    @Override
    public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
        if (!recordBinder.withinRange(sheetIndex, rowNum, columnNum)) {
            return;
        }

        recordBinder.setValue(currentRecord, columnNum, cellValue);
    }

    protected E createRecord() {
        if (constructor != null) {
            try {
                return constructor.newInstance();
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
