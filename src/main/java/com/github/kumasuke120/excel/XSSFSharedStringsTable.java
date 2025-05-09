package com.github.kumasuke120.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import java.io.Closeable;
import java.io.IOException;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;

/**
 * A wrapper class for the SharedStringsTable class in Apache POI.
 * This class provides compatibility with multiple versions of Apache POI (starting from 3.17).
 */
@ApiStatus.Internal
class XSSFSharedStringsTable implements Closeable {

    private static final MethodHandle getSharedStringsTableHandle = getGetSharedStringsTableHandle();

    private final Object table;
    private final MethodHandle getEntryAtHandle;
    private final MethodHandle getItemAtHandle;
    private final MethodHandle closeHandle;

    private XSSFSharedStringsTable(Object table) {
        this.table = table;
        this.getEntryAtHandle = getGetEntryAtHandle();
        this.getItemAtHandle = getGetItemAtHandle();
        this.closeHandle = getCloseHandle();
    }

    /**
     * Retrieves the shared strings table from the given XSSFReader.
     *
     * @param reader The XSSFReader instance.
     * @return An XSSFSharedStringsTable instance wrapping the shared strings table.
     * @throws IOException            If an I/O error occurs.
     * @throws InvalidFormatException If the format of the shared strings table is invalid.
     */
    public static XSSFSharedStringsTable getSharedStringsTable(XSSFReader reader) throws IOException, InvalidFormatException {
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
    public RichTextString getItemAt(int idx) {
        if (getItemAtHandle != null) {
            try {
                return (RichTextString) getItemAtHandle.invoke(table, idx);
            } catch (Throwable e) {
                if (e instanceof RuntimeException) {
                    throw (RuntimeException) e;
                } else {
                    throw new AssertionError(e);
                }
            }
        } else if (getEntryAtHandle != null) {
            CTRst entry;
            try {
                entry = (CTRst) getEntryAtHandle.invoke(table, idx);
            } catch (Throwable e) {
                if (e instanceof RuntimeException) {
                    throw (RuntimeException) e;
                } else {
                    throw new AssertionError(e);
                }
            }
            return new XSSFRichTextString(entry);
        } else {
            throw new IndexOutOfBoundsException("Index out of bounds: " + idx);
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

    private static @Nullable MethodHandle getGetSharedStringsTableHandle() {
        final Class<?> sharedStringsClass = getSharedStringsClass();
        final Class<?> sharedStringsTableClass = getSharedStringsTableClass();

        MethodHandle handle = null;
        // for Apache POI 5.x
        if (sharedStringsClass != null) {
            try {
                handle = MethodHandles.lookup().findVirtual(XSSFReader.class, "getSharedStringsTable",
                        MethodType.methodType(sharedStringsClass));
            } catch (NoSuchMethodException | IllegalAccessException ignore) {
            }
        }

        // for Apache POI 3.x and 4.x
        if (handle == null && sharedStringsTableClass != null) {
            try {
                handle = MethodHandles.lookup().findVirtual(XSSFReader.class, "getSharedStringsTable",
                        MethodType.methodType(sharedStringsTableClass));
            } catch (NoSuchMethodException | IllegalAccessException ignore) {
            }
        }

        return handle;
    }

    private static Class<?> getSharedStringsTableClass() {
        try {
            return Class.forName("org.apache.poi.xssf.model.SharedStringsTable");
        } catch (ClassNotFoundException e) {
            return null;
        }
    }

    private static Class<?> getSharedStringsClass() {
        try {
            return Class.forName("org.apache.poi.xssf.model.SharedStrings");
        } catch (ClassNotFoundException e) {
            return null;
        }
    }

    private MethodHandle getGetEntryAtHandle() {
        try {
            return MethodHandles.lookup().findVirtual(table.getClass(), "getEntryAt",
                    MethodType.methodType(CTRst.class, int.class));
        } catch (NoSuchMethodException | IllegalAccessException e) {
            return null;
        }
    }

    private MethodHandle getGetItemAtHandle() {
        try {
            return MethodHandles.lookup().findVirtual(table.getClass(), "getItemAt",
                    MethodType.methodType(RichTextString.class, int.class));
        } catch (NoSuchMethodException | IllegalAccessException e) {
            return null;
        }
    }

    private MethodHandle getCloseHandle() {
        try {
            return MethodHandles.lookup().findVirtual(table.getClass(), "close",
                    MethodType.methodType(void.class));
        } catch (NoSuchMethodException | IllegalAccessException e) {
            return null;
        }
    }

}
