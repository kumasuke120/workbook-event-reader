package com.github.kumasuke120.excel;

import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.Closeable;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

/**
 * A reader that reads Workbook in an event manner and it can only deal with values in cell (not charts)<br>
 */
public interface WorkbookEventReader extends Closeable {

    /**
     * Opens the specified file with an appropriate {@link WorkbookEventReader} if possible.
     *
     * @param filePath path of the file to be opened
     * @return {@link WorkbookEventReader} to read the specified file
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    @NotNull
    static WorkbookEventReader open(@NotNull(exception = NullPointerException.class) Path filePath) {
        return open(filePath, null);
    }

    /**
     * Opens the specified encrypted file with the given password and an appropriate
     * {@link WorkbookEventReader} if possible.
     *
     * @param filePath path of the file to be opened
     * @param password password to open the file
     * @return {@link WorkbookEventReader} to read the specified file
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    @NotNull
    static WorkbookEventReader open(@NotNull(exception = NullPointerException.class) Path filePath,
                                    @Nullable String password) {
        return new WorkbookAutoOpener(filePath, password).open();
    }

    /**
     * Opens the specified {@link InputStream} with an appropriate {@link WorkbookEventReader} if possible.
     *
     * @param in {@link InputStream} of the workbook to be opened
     * @return {@link WorkbookEventReader} to read the specified file
     * @throws NullPointerException <code>in</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    @NotNull
    static WorkbookEventReader open(@NotNull(exception = NullPointerException.class) InputStream in) {
        return open(in, null);
    }

    /**
     * Opens the specified encrypted {@link InputStream} with the given password and an appropriate
     * {@link WorkbookEventReader} if possible.
     *
     * @param in       {@link InputStream} of the workbook to be opened
     * @param password password to open the file
     * @return {@link WorkbookEventReader} to read the specified file
     * @throws NullPointerException <code>in</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    @NotNull
    static WorkbookEventReader open(@NotNull(exception = NullPointerException.class) InputStream in,
                                    @Nullable String password) {
        return new WorkbookAutoOpener(in, password).open();
    }

    /**
     * Starts to read the workbook through event handling, triggering events on the {@link EventHandler}
     * in a reasonable and recursive order: Document, Sheet, Row and Cell.<br>
     * This method can be called multiple times as long as this {@link WorkbookEventReader} is not closed.
     *
     * @param handler a {@link EventHandler} that handles read events as reading process going
     * @throws NullPointerException        <code>handler</code> is <code>null</code>
     * @throws WorkbookProcessException    errors happened when reading
     * @throws IllegalReaderStateException this {@link WorkbookEventReader} has been closed;
     *                                     called during reading process
     */
    void read(@NotNull(exception = NullPointerException.class) EventHandler handler);

    /**
     * Cancels reading process that is currently performing as soon as possible.
     * It cannot cancel the process immediately, but it will cancel the process before next event's happening.<br>
     * The event {@link EventHandler#onReadCancelled() onReadCancelled()} will be triggered when the
     * reading process has been cancelled successfully.
     *
     * @throws IllegalReaderStateException this {@link WorkbookEventReader} has been closed;
     *                                     called not during reading process
     */
    void cancel();

    /**
     * Closes this {@link WorkbookEventReader}.
     *
     * @throws WorkbookIOException errors happened when closing
     */
    void close();

    /**
     * An <code>EventHandler</code> that deals with event triggered during reading process of
     * {@link WorkbookEventReader}
     */
    interface EventHandler {
        /**
         * Get triggered when the document starts.
         */
        default void onStartDocument() {
            // no-op
        }

        /**
         * Get triggered when the document ends.
         */
        default void onEndDocument() {
            // no-op
        }

        /**
         * Get triggered whenever a sheet starts.
         *
         * @param sheetIndex the index of the sheet, starts with 0
         * @param sheetName  the name of the sheet
         */
        default void onStartSheet(int sheetIndex, @NotNull String sheetName) {
            // no-op
        }

        /**
         * Get triggered whenever a sheet ends.
         *
         * @param sheetIndex the index of the sheet, starts with 0
         */
        default void onEndSheet(int sheetIndex) {
            // no-op
        }

        /**
         * Get triggered whenever a row starts.
         *
         * @param sheetIndex the index of the containing sheet, starts with 0
         * @param rowNum     the index of the row, starts with 0
         */
        default void onStartRow(int sheetIndex, int rowNum) {
            // no-op
        }

        /**
         * Get triggered whenever a row ends.
         *
         * @param sheetIndex the index of the containing sheet, starts with 0
         * @param rowNum     the index of the row, starts with 0
         */
        default void onEndRow(int sheetIndex, int rowNum) {
            // no-op
        }

        /**
         * Get triggered whenever a cell and its value could be handled.<br>
         * The value of the cell will be of nine different types (including <code>null</code>)
         * as follows:<br>
         * <table>
         * <caption>All Possible Types And Their Descriptions</caption>
         * <thead>
         * <tr>
         * <th>Possible Type</th>
         * <th>Description</th>
         * </tr>
         * </thead>
         * <tbody>
         * <tr>
         * <td><code>null</code></td>
         * <td>blank or error value</td>
         * </tr>
         * <tr>
         * <td>{@link Boolean}</td>
         * <td><code>true</code> or <code>false</code>calculated by formula</td>
         * </tr>
         * <tr>
         * <td>{@link Integer}</td>
         * <td>numeric value that could be treat as <code>int</code></td>
         * </tr>
         * <tr>
         * <td>{@link Long}</td>
         * <td>numeric value that could be treat as <code>long</code></td>
         * </tr>
         * <tr>
         * <td>{@link Double}</td>
         * <td>numeric value that could be treat as <code>double</code></td>
         * </tr>
         * <tr>
         * <td>{@link String}</td>
         * <td>string value, or value from a cell that marked as text</td>
         * </tr>
         * <tr>
         * <td>{@link LocalTime}</td>
         * <td>time value whose containing cell is a non-text cell</td>
         * </tr>
         * <tr>
         * <td>{@link LocalDate}</td>
         * <td>date value without a time part whose containing cell is a non-text cell</td>
         * </tr>
         * <tr>
         * <td>{@link LocalDateTime}</td>
         * <td>date-time value whose containing cell is a non-text cell</td>
         * </tr>
         * </tbody>
         * </table>
         * The value will be wrapped in a {@link CellValue}, you could use {@link CellValue#originalValue()} to
         * get its original value or get the value of the type you want easily from the wrapped object.
         *
         * @param sheetIndex the index of the containing sheet, starts with 0
         * @param rowNum     the index of the containing row, starts with 0
         * @param columnNum  the index of the cell, starts with 0
         * @param cellValue  the value read and converted from the corresponding cell, it won't be <code>null</code>
         */
        default void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
            // no-op
        }

        /**
         * Get triggered when the reading process actually being cancelled after the invocation of the
         * {@link WorkbookEventReader#cancel()} method.
         */
        default void onReadCancelled() {
            // no-op
        }
    }

}
