package com.github.kumasuke120.excel.handler;

import com.github.kumasuke120.excel.CellValue;
import com.github.kumasuke120.excel.WorkbookEventReader;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

@ApiStatus.Experimental
public class WorkbookRecordExtractor<E> implements WorkbookEventReader.EventHandler {

    private final WorkbookEventReader reader;

    private final Class<E> recordClass;

    private final WorkbookRecordBinder<E> recordBinder;

    private final Constructor<E> constructor;

    private List<E> result;

    private E currentRecord;

    public WorkbookRecordExtractor(@NotNull(exception = NullPointerException.class) WorkbookEventReader reader,
                                   @NotNull(exception = NullPointerException.class) Class<E> recordClass) {

        this.reader = Objects.requireNonNull(reader);
        this.recordClass = HandlerUtils.ensureWorkbookRecordClass(recordClass);
        this.recordBinder = new WorkbookRecordBinder<>(recordClass);
        this.constructor = findConstructor();
    }

    private Constructor<E> findConstructor() {
        try {
            return recordClass.getConstructor();
        } catch (NoSuchMethodException e) {
            return null;
        }
    }

    public List<E> getResult() {
        final List<E> ret = new ArrayList<>(result);
        result = null;
        return ret;
    }

    @Override
    public void onStartDocument() {
        result = new ArrayList<>();
    }

    @Override
    public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
/*        if (!recordBinder.withinRange(sheetIndex)) {
            reader.cancel();
        }*/
    }

    @Override
    public void onEndSheet(int sheetIndex) {
/*        if (!recordBinder.withinRange(sheetIndex)) {
            reader.cancel();
        }*/
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
