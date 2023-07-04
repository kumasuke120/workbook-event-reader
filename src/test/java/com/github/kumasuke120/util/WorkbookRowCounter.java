package com.github.kumasuke120.util;

import com.github.kumasuke120.excel.WorkbookEventReader;
import org.jetbrains.annotations.NotNull;

import java.util.HashMap;
import java.util.Map;

public class WorkbookRowCounter implements WorkbookEventReader.EventHandler {

    private final Map<Integer, Integer> sheetToRows = new HashMap<>();

    @Override
    public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
        sheetToRows.put(sheetIndex, 0);
    }

    @Override
    public void onEndRow(int sheetIndex, int rowNum) {
        sheetToRows.merge(sheetIndex, 1, Integer::sum);
    }

    public int getSheetCount() {
        return sheetToRows.keySet().stream()
                .max(Integer::compareTo)
                .map(count -> count + 1)
                .orElse(0);
    }

    public int getRowCount(int sheetIndex) {
        final Integer rows = sheetToRows.get(sheetIndex);
        return rows == null ? 0 : rows;
    }

}
