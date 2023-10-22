package com.github.kumasuke120.excel;

import com.github.kumasuke120.excel.WorkbookAutoOpener.ReaderConstructor;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

class WorkbookAutoOpenerTest {

    @Test
    void constructor_newInstance() {
        ReaderConstructor constructor = new ReaderConstructor(NormalWorkbookReader.class);
        assertDoesNotThrow(() -> {
            final Path tempFile = Files.createTempFile(null, ".xlsx");
            try (final WorkbookEventReader r = constructor.newInstance(tempFile, null)) {
                assertNotNull(r);
            } finally {
                try {
                    Files.delete(tempFile);
                } catch (IOException ignored) {
                }
            }


            try (final ByteArrayInputStream in = new ByteArrayInputStream(new byte[0]);
                 final WorkbookEventReader r = constructor.newInstance(in, null)) {
                assertNotNull(r);
            }
        });

        ReaderConstructor constructor2 = new ReaderConstructor(OpenErrorWorkbookReader.class);
        assertThrows(Error.class, () -> {
            final Path tempFile = Files.createTempFile(null, ".xlsx");
            try (final WorkbookEventReader r = constructor2.newInstance(tempFile, null)) {
                assertNull(r);
            } finally {
                try {
                    Files.delete(tempFile);
                } catch (IOException ignored) {
                }
            }


            try (final ByteArrayInputStream in = new ByteArrayInputStream(new byte[0]);
                 final WorkbookEventReader r = constructor2.newInstance(in, null)) {
                assertNull(r);
            }
        });

        ReaderConstructor constructor3 = new ReaderConstructor(OpenWorkbookIOExceptionWorkbookReader.class);
        assertThrows(WorkbookIOException.class, () -> {
            final Path tempFile = Files.createTempFile(null, ".xlsx");
            try (final WorkbookEventReader r = constructor3.newInstance(tempFile, null)) {
                assertNull(r);
            } finally {
                try {
                    Files.delete(tempFile);
                } catch (IOException ignored) {
                }
            }


            try (final ByteArrayInputStream in = new ByteArrayInputStream(new byte[0]);
                 final WorkbookEventReader r = constructor3.newInstance(in, null)) {
                assertNull(r);
            }
        });

        ReaderConstructor constructor4 = new ReaderConstructor(OpenRuntimeExceptionWorkbookReader.class);
        assertThrows(WorkbookIOException.class, () -> {
            final Path tempFile = Files.createTempFile(null, ".xlsx");
            try (final WorkbookEventReader r = constructor4.newInstance(tempFile, null)) {
                assertNull(r);
            } finally {
                try {
                    Files.delete(tempFile);
                } catch (IOException ignored) {
                }
            }


            try (final ByteArrayInputStream in = new ByteArrayInputStream(new byte[0]);
                 final WorkbookEventReader r = constructor4.newInstance(in, null)) {
                assertNull(r);
            }
        });
    }

    @SuppressWarnings("unused")
    private static class OpenRuntimeExceptionWorkbookReader extends AbstractTestWorkbookReader {

        public OpenRuntimeExceptionWorkbookReader(@NotNull(exception = NullPointerException.class) InputStream in,
                                                  @Nullable String password) {
            super(in, password);
        }

        public OpenRuntimeExceptionWorkbookReader(@NotNull(exception = NullPointerException.class) Path filePath,
                                                  @Nullable String password) {
            super(filePath, password);
        }

        @Override
        void doOpen(@NotNull InputStream in, @Nullable String password) {
            throw new RuntimeException("InputStream");
        }

        @Override
        void doOpen(@NotNull Path filePath, @Nullable String password) {
            throw new RuntimeException("Path");
        }

    }

    @SuppressWarnings("unused")
    private static class OpenWorkbookIOExceptionWorkbookReader extends AbstractTestWorkbookReader {

        public OpenWorkbookIOExceptionWorkbookReader(@NotNull(exception = NullPointerException.class) InputStream in,
                                                     @Nullable String password) {
            super(in, password);
        }

        public OpenWorkbookIOExceptionWorkbookReader(@NotNull(exception = NullPointerException.class) Path filePath,
                                                     @Nullable String password) {
            super(filePath, password);
        }

        @Override
        void doOpen(@NotNull InputStream in, @Nullable String password) {
            throw new WorkbookIOException("InputStream", new IOException());
        }

        @Override
        void doOpen(@NotNull Path filePath, @Nullable String password) {
            throw new WorkbookIOException("Path", new IOException());
        }

    }


    @SuppressWarnings("unused")
    private static class OpenErrorWorkbookReader extends AbstractTestWorkbookReader {

        public OpenErrorWorkbookReader(@NotNull(exception = NullPointerException.class) InputStream in,
                                       @Nullable String password) {
            super(in, password);
        }

        public OpenErrorWorkbookReader(@NotNull(exception = NullPointerException.class) Path filePath,
                                       @Nullable String password) {
            super(filePath, password);
        }

        @Override
        void doOpen(@NotNull InputStream in, @Nullable String password) {
            throw new AssertionError("InputStream");
        }

        @Override
        void doOpen(@NotNull Path filePath, @Nullable String password) {
            throw new AssertionError("Path");
        }

    }


    @SuppressWarnings("unused")
    private static class NormalWorkbookReader extends AbstractTestWorkbookReader {

        public NormalWorkbookReader(@NotNull(exception = NullPointerException.class) InputStream in,
                                    @Nullable String password) {
            super(in, password);
        }

        public NormalWorkbookReader(@NotNull(exception = NullPointerException.class) Path filePath,
                                    @Nullable String password) {
            super(filePath, password);
        }

        @Override
        void doOpen(@NotNull InputStream in, @Nullable String password) {

        }

        @Override
        void doOpen(@NotNull Path filePath, @Nullable String password) {

        }

    }

    @SuppressWarnings("unused")
    private abstract static class AbstractTestWorkbookReader extends AbstractWorkbookEventReader {

        public AbstractTestWorkbookReader(@NotNull(exception = NullPointerException.class) InputStream in,
                                          @Nullable String password) {
            super(in, password);
        }

        public AbstractTestWorkbookReader(@NotNull(exception = NullPointerException.class) Path filePath,
                                          @Nullable String password) {
            super(filePath, password);
        }

        @Override
        final @NotNull ReaderCleanAction createCleanAction() {
            return new ReaderCleanAction() {
                @Override
                void doClean() {

                }
            };
        }

        @Override
        final void doRead(@NotNull EventHandler handler) {

        }

    }

}