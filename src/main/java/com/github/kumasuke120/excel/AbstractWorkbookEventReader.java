package com.github.kumasuke120.excel;

import org.jetbrains.annotations.Contract;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Objects;

/**
 * The base class for {@link WorkbookEventReader}, containing common methods and utilities
 */
abstract class AbstractWorkbookEventReader implements WorkbookEventReader {

    private static final ThreadLocal<WorkbookEventReader> currentReader = new ThreadLocal<>();

    private final ReaderCleanAction cleanAction;

    private volatile boolean closed = false;
    private volatile boolean reading = false;

    /**
     * Creates a new {@link AbstractWorkbookEventReader} based on the given file {@link InputStream}
     * and the given password if possible.
     *
     * @param in       {@link InputStream} of the workbook to be read
     * @param password password to open the file
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    AbstractWorkbookEventReader(@NotNull(exception = NullPointerException.class) InputStream in,
                                @Nullable String password) {
        Objects.requireNonNull(in);

        doOnStartOpen();
        try {
            doOpen(in, password);
        } catch (Exception e) {
            throw new WorkbookIOException("Cannot open workbook", e);
        }

        cleanAction = createCleanAction();
    }

    /**
     * Creates a new {@link AbstractWorkbookEventReader} based on the given file path
     * and the given password if possible.
     *
     * @param filePath {@link Path} of the workbook to be read
     * @param password password to open the file
     * @throws NullPointerException <code>filePath</code> is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    AbstractWorkbookEventReader(@NotNull(exception = NullPointerException.class) Path filePath,
                                @Nullable String password) {
        Objects.requireNonNull(filePath);

        doOnStartOpen();
        try {
            doOpen(filePath, password);
        } catch (Exception e) {
            throw new WorkbookIOException("Cannot open workbook", e);
        }

        cleanAction = createCleanAction();
    }

    /**
     * Closes the {@link Closeable}. Exception thrown during closing will be suppressed and
     * add to the previous caught {@link Exception} if necessary. Otherwise, the exception will
     * be thrown.
     *
     * @param closeable object to be closed
     * @param caught    previous caught {@link Exception}
     * @throws Exception previous caught {@link Exception} or {@link Exception} thrown when closing
     */
    static void suppressClose(@Nullable Closeable closeable, @Nullable Exception caught) throws Exception {
        Exception thrown = caught;
        try {
            if (closeable != null) {
                closeable.close();
            }
        } catch (IOException e) {
            if (thrown != null) {
                e.addSuppressed(thrown);
            }
            thrown = e;
        }

        if (thrown != null) throw thrown;
    }

    /**
     * Callback which will be called before the actual opening process is completed or cancelled.<br>
     * <br>
     * * This method shouldn't throw any kind of exceptions, including unchecked exceptions.
     */
    void doOnStartOpen() {
        // no-op
    }

    /**
     * Opens the specified file with this {@link AbstractWorkbookEventReader}.<br>
     * <br>
     * * This method should not throw any {@link WorkbookEventReaderException}, because it is the duty of its
     * caller to wrap every exception it throws into {@link WorkbookEventReaderException}.
     *
     * @param in       {@link InputStream} of the workbook to be opened, and it won't be <code>null</code>
     * @param password password to open the file, and it might be <code>null</code>
     * @throws Exception any exception occurred during opening process
     */
    abstract void doOpen(@NotNull InputStream in, @Nullable String password) throws Exception;

    /**
     * Opens the specified file with this {@link AbstractWorkbookEventReader}.<br>
     * <br>
     * * This method should not throw any {@link WorkbookEventReaderException}, because it is the duty of its
     * caller to wrap every exception it throws into {@link WorkbookEventReaderException}.
     *
     * @param filePath {@link Path} of the workbook to be opened, and it won't be <code>null</code>
     * @param password password to open the file, and it might be <code>null</code>
     * @throws Exception any exception occurred during opening process
     */
    abstract void doOpen(@NotNull Path filePath, @Nullable String password) throws Exception;

    /**
     * Creates a resource-cleaning action to close all resources this {@link WorkbookEventReader}
     * has opened.<br>
     * <br>
     * * The implementation of {@link ReaderCleanAction} should not be an anonymous or non-static inner class.
     * In addition, it should not contain any reference of the containing {@link WorkbookEventReader}.
     *
     * @return a non-anonymous instance of {@link ReaderCleanAction}
     */
    @NotNull
    @Contract("-> new")
    abstract ReaderCleanAction createCleanAction();

    /**
     * {@inheritDoc}
     */
    @Override
    public final void read(@NotNull(exception = NullPointerException.class) EventHandler handler) {
        assertNotClosed();
        assertNotBeingRead();

        Objects.requireNonNull(handler);

        final EventHandler delegate = new CancelFastEventHandler(handler);

        reading = true;
        currentReader.set(this);

        try {
            doRead(delegate);
        } catch (CancelReadingException ignored) {
            // stops parsing and cancels reading
        } catch (Exception e) {
            if (e instanceof WorkbookEventReaderException) {
                throw (WorkbookEventReaderException) e;
            } else {
                throw new WorkbookProcessException(e);
            }
        } finally {
            currentReader.remove();
            reading = false;
        }
    }

    /**
     * Starts to read the workbook through event handling, triggering events on the {@link EventHandler}
     * in a reasonable and recursive order: Document, Sheet, Row and Cell.<br>
     * <br>
     * * This method should be able to called multiple times as long as this {@link AbstractWorkbookEventReader}
     * is not closed.<br>
     * * This method or its user may throw a {@link WorkbookEventReaderException} which will be re-thrown in
     * {@link #read(EventHandler)}.<br>
     * * Normally, this method does not have to process cancellation triggered by user.
     *
     * @param handler an non-<code>null</code> {@link EventHandler} that handles read events as reading process going
     * @throws Exception any exception occurred during reading process
     */
    abstract void doRead(@NotNull EventHandler handler) throws Exception;

    /**
     * {@inheritDoc}
     */
    @Override
    public final void cancel() {
        assertNotClosed();
        assertBeingRead();

        reading = false;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public final void close() {
        if (!closed) {
            reading = false;
            closed = true;

            cleanAction.run();
        }
    }

    /**
     * Asserts the reader is not being closed. Otherwise, it throws {@link IllegalReaderStateException}.
     *
     * @throws IllegalReaderStateException the reader has been closed
     */
    private void assertNotClosed() {
        if (closed) {
            throw new IllegalReaderStateException("This '" + getClass().getSimpleName() + "' has been closed");
        }
    }

    /**
     * Asserts the reader is not being read. Otherwise, it throws {@link IllegalReaderStateException}.
     *
     * @throws IllegalReaderStateException the reader is being read
     */
    private void assertNotBeingRead() {
        if (reading) {
            throw new IllegalReaderStateException("This '" + getClass().getSimpleName() + "' is being read");
        }
    }

    /**
     * Asserts the reader is being read. Otherwise, it throws {@link IllegalReaderStateException}.
     *
     * @throws IllegalReaderStateException the reader is not being read
     */
    private void assertBeingRead() {
        if (!reading) {
            throw new IllegalReaderStateException("This '" + getClass().getSimpleName() + "' is not being read");
        }
    }

    /**
     * A clean action for closes the resources opened by a {@link WorkbookEventReader}
     */
    static abstract class ReaderCleanAction implements Runnable {
        private volatile boolean cleaned = false;

        @Override
        public final void run() {
            if (!cleaned) {
                try {
                    doClean();
                } catch (Exception e) {
                    throw new WorkbookIOException("Exception encountered when closing the workbook file", e);
                } finally {
                    cleaned = true;
                }
            }
        }

        /**
         * Closes the resources from the related {@link WorkbookEventReader}. <br>
         * <br>
         * * This method should not throw any {@link WorkbookEventReaderException}, because it is the duty of its
         * caller to wrap every exception it throws into {@link WorkbookEventReaderException}.
         *
         * @throws Exception any exception occurred during closing process
         */
        abstract void doClean() throws Exception;
    }

    /**
     * A <code>RuntimeException</code> that stops EventHandler, not an actual exception
     */
    private static class CancelReadingException extends RuntimeException {
    }

    /**
     * An <code>EventHandler</code> that checks reading state before any event is triggered
     */
    private class CancelFastEventHandler implements EventHandler {
        private final EventHandler handler;
        private volatile boolean cancelled;

        private CancelFastEventHandler(@NotNull EventHandler handler) {
            this.handler = handler;
            this.cancelled = false;
        }

        @Override
        public void onStartDocument() {
            onEvent(handler::onStartDocument);
        }

        @Override
        public void onEndDocument() {
            onEvent(handler::onEndDocument);
        }

        @Override
        public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
            onEvent(() -> handler.onStartSheet(sheetIndex, sheetName));
        }

        @Override
        public void onEndSheet(int sheetIndex) {
            onEvent(() -> handler.onEndSheet(sheetIndex));
        }

        @Override
        public void onStartRow(int sheetIndex, int rowNum) {
            onEvent(() -> handler.onStartRow(sheetIndex, rowNum));
        }

        @Override
        public void onEndRow(int sheetIndex, int rowNum) {
            onEvent(() -> handler.onEndRow(sheetIndex, rowNum));
        }

        @Override
        public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
            onEvent(() -> handler.onHandleCell(sheetIndex, rowNum, columnNum, cellValue));
        }

        private void onEvent(Runnable runnable) {
            if (reading) {
                runnable.run();
            } else {
                doOnCancelled();
            }
        }

        @Override
        public void onReadCancelled() {
            if (!cancelled) {
                cancelled = true;
                handler.onReadCancelled();
            }
        }

        private void doOnCancelled() {
            // triggers when the reading process was cancelled
            onReadCancelled();

            // throws an exception to abort processing
            throw new CancelReadingException();
        }

    }

    static final class ReadContextImpl implements WorkbookEventReader.ReadContext {

        private final WorkbookEventReader reader;

        ReadContextImpl(@NotNull WorkbookEventReader reader) {
            this.reader = reader;
        }

        /**
         * Returns the current reading context based on current reader of the thread.
         *
         * @return current reading context if any
         * @throws IllegalReaderStateException cannot get the current reader of the thread
         */
        static ReadContextImpl current() {
            final WorkbookEventReader reader = currentReader.get();
            if (reader == null) {
                throw new IllegalReaderStateException("No WorkbookEventReader is being read");
            }

            return new ReadContextImpl(reader);
        }

        /**
         * {@inheritDoc}
         */
        @Override
        public void cancel() {
            reader.cancel();
        }

    }

}
