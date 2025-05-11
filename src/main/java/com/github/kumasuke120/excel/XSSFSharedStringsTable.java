package com.github.kumasuke120.excel;

import com.github.kumasuke120.excel.util.ReflectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import java.io.Closeable;
import java.io.IOException;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodType;

/**
 * A wrapper class for the SharedStringsTable class in Apache POI.
 * This class provides compatibility with multiple versions of Apache POI (starting from 3.17).
 */
@ApiStatus.Internal
class XSSFSharedStringsTable implements Closeable {

    private static final String SHARED_STRINGS_CLASS_NAME = "org.apache.poi.xssf.model.SharedStrings";
    private static final String SHARED_STRINGS_TABLE_CLASS_NAME = "org.apache.poi.xssf.model.SharedStringsTable";
    private static final MethodHandle getSharedStringsTableHandle = getGetSharedStringsTableHandle();

    private final Object table;
    private final MethodHandle getEntryAtHandle;
    private final MethodHandle getItemAtHandle;
    private final MethodHandle closeHandle;

    private XSSFSharedStringsTable(Object table) {
        this.table = table;
        this.getEntryAtHandle = ReflectionUtils.findVirtualHandle(table.getClass(), "getEntryAt",
                MethodType.methodType(CTRst.class, int.class));
        this.getItemAtHandle = ReflectionUtils.findVirtualHandle(table.getClass(), "getItemAt",
                MethodType.methodType(RichTextString.class, int.class));
        this.closeHandle = ReflectionUtils.findVirtualHandle(table.getClass(), "close",
                MethodType.methodType(void.class));
    }

    /**
     * Retrieves the shared strings table from the given XSSFReader.
     *
     * @param reader The XSSFReader instance.
     * @return An XSSFSharedStringsTable instance wrapping the shared strings table.
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the format of the shared strings table is invalid.
     */
    public static XSSFSharedStringsTable getSharedStringsTable(@NotNull XSSFReader reader) throws IOException, InvalidFormatException {
        final Object table = getSharedStringsTable0(reader);
        if (table == null) {
            return null;
        } else {
            return new XSSFSharedStringsTable(table);
        }
    }

    /**
     * Return a string item by index
     *
     * @param idx index of item to return.
     * @return the item at the specified position in this Shared String table.
     */
    @NotNull
    public RichTextString getItemAt(int idx) {
        if (getItemAtHandle != null) {
            return invokeMethodHandleForItemAt(getItemAtHandle, table, idx);
        } else if (getEntryAtHandle != null) {
            final CTRst entry = invokeMethodHandleForItemAt(getEntryAtHandle, table, idx);
            return new XSSFRichTextString(entry);
        } else {
            throw new IndexOutOfBoundsException("Index out of bounds: " + idx);
        }
    }

    @SuppressWarnings("unchecked")
    private static <T> T invokeMethodHandleForItemAt(MethodHandle handle, Object... args) {
        try {
            return (T) handle.invokeWithArguments(args);
        } catch (Throwable e) {
            if (e instanceof RuntimeException) {
                throw (RuntimeException) e;
            } else {
                throw new AssertionError(e);
            }
        }
    }

    /**
     * Closes the shared strings table.
     *
     * @throws IOException If an error occurs while closing the table.
     */
    @Override
    public void close() throws IOException {
        if (closeHandle == null) {
            return;
        }

        try {
            closeHandle.invoke(table);
        } catch (Throwable e) {
            if (e instanceof IOException) {
                throw (IOException) e;
            } else {
                throw new AssertionError(e);
            }
        }
    }

    private static Object getSharedStringsTable0(XSSFReader reader) throws IOException, InvalidFormatException {
        if (getSharedStringsTableHandle == null) {
            throw new UnsupportedOperationException("cannot find getSharedStringsTable() method for XSSFReader");
        }

        try {
            return getSharedStringsTableHandle.invoke(reader);
        } catch (Throwable e) {
            if (e instanceof IOException) {
                throw (IOException) e;
            } else if (e instanceof InvalidFormatException) {
                throw (InvalidFormatException) e;
            } else {
                throw new AssertionError(e);
            }
        }
    }

    private static MethodHandle getGetSharedStringsTableHandle() {
        final Class<?> sharedStringsClass = ReflectionUtils.getClass(SHARED_STRINGS_CLASS_NAME);
        final Class<?> sharedStringsTableClass = ReflectionUtils.getClass(SHARED_STRINGS_TABLE_CLASS_NAME);

        MethodHandle handle = null;

        // for Apache POI 5.x
        if (sharedStringsClass != null) {
            handle = ReflectionUtils.findVirtualHandle(XSSFReader.class, "getSharedStringsTable",
                    MethodType.methodType(sharedStringsClass));
        }

        // for Apache POI 3.x and 4.x
        if (handle == null && sharedStringsTableClass != null) {
            handle = ReflectionUtils.findVirtualHandle(XSSFReader.class, "getSharedStringsTable",
                    MethodType.methodType(sharedStringsTableClass));
        }

        return handle;
    }

}
