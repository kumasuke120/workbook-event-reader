package com.github.kumasuke120.excel;

import com.github.kumasuke120.util.LightWeightConstructor;
import com.github.kumasuke120.util.ResourceUtil;
import com.github.kumasuke120.util.WorkbookRowCounter;
import com.github.kumasuke120.util.XmlUtil;
import org.apache.poi.EncryptedDocumentException;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.junit.jupiter.api.Assertions;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Stack;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

abstract class AbstractWorkbookEventReaderTest<R extends AbstractWorkbookEventReader> {

    private final String normalFileName;
    private final String encryptedFileName;
    private final Class<R> readerClass;

    private String sampleReadFileName = "sample-output.xml";
    private String sampleCancelFileName = "sample-output-2.xml";

    AbstractWorkbookEventReaderTest(@NotNull String normalFileName,
                                    @Nullable String encryptedFileName,
                                    @NotNull Class<R> readerClass) {
        this.normalFileName = normalFileName;
        this.encryptedFileName = encryptedFileName;
        this.readerClass = readerClass;
    }

    @SuppressWarnings("SameParameterValue")
    void setSampleReadFileName(@NotNull String sampleReadFileName) {
        this.sampleReadFileName = sampleReadFileName;
    }

    @SuppressWarnings("SameParameterValue")
    void setSampleCancelFileName(@NotNull String sampleCancelFileName) {
        this.sampleCancelFileName = sampleCancelFileName;
    }

    @SuppressWarnings("EmptyTryBlock")
    final void dealWithReader(@NotNull Consumer<WorkbookEventReader> consumer) {
        // region constructor(Path)
        final Path filePath = ResourceUtil.getPathOfClasspathResource(normalFileName);
        try (final WorkbookEventReader reader = pathConstructor().newInstance(filePath)) {
            consumer.accept(reader);
        }
        assertThrows(NullPointerException.class, () -> {
            try (final WorkbookEventReader ignore = pathConstructor().newInstance((Object) null)) {
                // no-op
            }
        });
        // endregion

        // region constructor(InputStream)
        try (final InputStream in = ClassLoader.getSystemResourceAsStream(normalFileName)) {
            try (final WorkbookEventReader reader = inputStreamConstructor().newInstance(in)) {
                consumer.accept(reader);
            }
        } catch (IOException e) {
            throw new AssertionError(e);
        }
        assertThrows(NullPointerException.class, () -> {
            try (final WorkbookEventReader ignore = inputStreamConstructor().newInstance((Object) null)) {
                // no-op
            }
        });
        // endregion

        if (encryptedFileName == null || "".equals(encryptedFileName)) { // skips encryption test
            return;
        }

        // region constructor(Path, String)
        final Path filePath2 = ResourceUtil.getPathOfClasspathResource(encryptedFileName);
        try (final WorkbookEventReader reader = pathAndPasswordConstructor()
                .newInstance(filePath2, WorkbookReaderTest.CORRECT_PASSWORD)) {
            consumer.accept(reader);
        }
        assertThrows(WorkbookIOException.class, () -> {
            try (final WorkbookEventReader ignore = pathAndPasswordConstructor()
                    .newInstance(filePath2, WorkbookReaderTest.randomWrongPassword())) {
                // no-op
            } catch (WorkbookIOException e) {
                assertTrue(e.getCause() instanceof EncryptedDocumentException);
                throw e;
            }
        });
        // endregion

        // region constructor(InputStream, String)
        try (final InputStream in = ClassLoader.getSystemResourceAsStream(encryptedFileName)) {
            try (final WorkbookEventReader reader = inputStreamAndPasswordConstructor()
                    .newInstance(in, WorkbookReaderTest.CORRECT_PASSWORD)) {
                consumer.accept(reader);
            }
        } catch (IOException e) {
            throw new AssertionError(e);
        }
        assertThrows(WorkbookIOException.class, () -> {
            try (final InputStream in = ClassLoader.getSystemResourceAsStream(encryptedFileName)) {
                try (final WorkbookEventReader ignore = inputStreamAndPasswordConstructor()
                        .newInstance(in, WorkbookReaderTest.randomWrongPassword())) {
                    // no-op
                } catch (WorkbookIOException e) {
                    assertTrue(e.getCause() instanceof EncryptedDocumentException);
                    throw e;
                }
            } catch (IOException e) {
                throw new AssertionError(e);
            }
        });
        // endregion
    }

    @NotNull
    LightWeightConstructor<R> pathConstructor() {
        return new LightWeightConstructor<>(readerClass, Path.class);
    }

    @NotNull
    private LightWeightConstructor<R> pathAndPasswordConstructor() {
        return new LightWeightConstructor<>(readerClass, Path.class, String.class);
    }

    @NotNull
    private LightWeightConstructor<R> inputStreamConstructor() {
        return new LightWeightConstructor<>(readerClass, InputStream.class);
    }

    @NotNull
    private LightWeightConstructor<R> inputStreamAndPasswordConstructor() {
        return new LightWeightConstructor<>(readerClass, InputStream.class, String.class);
    }

    void read() {
        dealWithReader(reader -> {
            final TestEventHandler handler = new TestEventHandler();
            reader.read(handler);

            final String xml = handler.getXml();
            assertSameWithSample(sampleReadFileName, xml);

            // read cannot be started when reader is in reading process
            reader.read(new TestEventHandler() {
                @Override
                public void onStartDocument() {
                    assertThrows(IllegalReaderStateException.class, () -> reader.read(handler));
                }
            });
        });

        dealWithReader(reader -> {
            final WorkbookEventReader.EventHandler handler = new WorkbookEventReader.EventHandler() {
                @Override
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
                    if (Integer.class.equals(cellValue.originalType())) {
                        // triggers an CellValueCastException
                        cellValue.localDateValue();
                    }
                }
            };

            assertThrows(WorkbookProcessException.class, () -> reader.read(handler));
        });

        dealWithReader(reader -> {
            final WorkbookEventReader.EventHandler handler = new WorkbookEventReader.EventHandler() {
                @Override
                public void onStartDocument() {
                    reader.read(new WorkbookEventReader.EventHandler() {
                    });
                }
            };

            assertThrows(IllegalReaderStateException.class, () -> reader.read(handler));
        });

        dealWithReader(reader -> {
            final WorkbookRowCounter counter1 = new WorkbookRowCounter();
            reader.read(counter1);

            final WorkbookRowCounter counter2 = new WorkbookRowCounter();
            reader.read(counter2);

            assertEquals(counter1.getSheetCount(), counter2.getSheetCount());
            for (int i = 0; i < counter1.getSheetCount(); i++) {
                final int rowCount1 = counter1.getRowCount(i);
                final int rowCount2 = counter2.getRowCount(i);

                assertEquals(rowCount1, rowCount2, "sheetIndex = " + i);
            }
        });
    }

    void cancel() {
        dealWithReader(reader -> {
            final boolean[] cancelledRef = {false};
            final TestEventHandler handler = new TestEventHandler() {
                @Override
                public void onEndSheet(int sheetIndex) {
                    super.onEndSheet(sheetIndex);

                    if (sheetIndex == 0 && reader != null) {
                        reader.cancel();
                    }
                }

                @Override
                public void onReadCancelled() {
                    cancelledRef[0] = true;
                }
            };

            reader.read(handler);

            handler.onEndDocument();
            final String xml = handler.getXml();
            assertSameWithSample(sampleCancelFileName, xml);

            // reader has been cancelled
            assertTrue(cancelledRef[0]);

            // reader is not being read
            assertThrows(IllegalReaderStateException.class, reader::cancel);
        });

        dealWithReader(reader -> {
            final boolean[] cancelledRef = {false};
            final WorkbookEventReader.EventHandler handler = new WorkbookEventReader.EventHandler() {
                @Override
                public void onStartRow(int sheetIndex, int rowNum) {
                    reader.cancel();
                }

                @Override
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
                    throw new AssertionError();
                }

                @Override
                public void onReadCancelled() {
                    cancelledRef[0] = true;
                }
            };

            reader.read(handler);

            assertTrue(cancelledRef[0]);
        });
    }

    void close() {
        dealWithReader(reader -> {
            reader.close();

            assertThrows(IllegalReaderStateException.class, () ->
                    reader.read(new TestEventHandler() {
                        @Override
                        public void onEndSheet(int sheetIndex) {
                            if (sheetIndex == 0) {
                                reader.close();
                            }
                            super.onEndSheet(sheetIndex);
                        }
                    }));

            assertThrows(IllegalReaderStateException.class, () -> reader.read(new TestEventHandler()));

            // at this point, the reader has been closed
            // however, at the end of dealWithReader, the close method will be invoked
            // one more time, that's intended for testing if reader could be closed multiple times
        });
    }

    private void assertSameWithSample(String sampleFileName, String actualXml) {
        final String expectedXml;
        try {
            expectedXml = ResourceUtil.loadClasspathResourceToString(sampleFileName);
        } catch (IOException e) {
            throw new AssertionError(e);
        }
        Assertions.assertTrue(XmlUtil.isSameXml(expectedXml, actualXml));
    }

    private static class TestEventHandler implements WorkbookEventReader.EventHandler {
        private final Stack<Integer> sheetStack;
        private final Stack<Integer> rowStack;
        private final StringBuilder xml;

        TestEventHandler() {
            xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            sheetStack = new Stack<>();
            rowStack = new Stack<>();
        }

        String getXml() {
            return xml.toString();
        }

        @Override
        public void onStartDocument() {
            xml.append("<document>");
        }

        @Override
        public void onEndDocument() {
            xml.append("</document>");
        }

        @Override
        public void onStartSheet(int sheetIndex, @NotNull String sheetName) {
            xml.append("<sheet index=\"")
                    .append(sheetIndex)
                    .append("\" name=\"")
                    .append(sheetName)
                    .append("\">");

            sheetStack.push(sheetIndex);

            if (sheetIndex == 0) {
                assertTrue("Sheet1".equals(sheetName) || "".equals(sheetName));
            } else if (sheetIndex == 1) {
                assertEquals("Sheet2", sheetName);
            } else {
                throw new AssertionError();
            }
        }

        @Override
        public void onEndSheet(int sheetIndex) {
            xml.append("</sheet>");

            final int poppedSheetIndex = sheetStack.pop();
            assertEquals(sheetIndex, poppedSheetIndex);
        }

        @Override
        public void onStartRow(int sheetIndex, int rowNum) {
            xml.append("<row index=\"")
                    .append(rowNum)
                    .append("\">");

            final int currentSheetIndex = sheetStack.peek();
            assertEquals(currentSheetIndex, sheetIndex);

            rowStack.push(rowNum);
        }

        @Override
        public void onEndRow(int sheetIndex, int rowNum) {
            xml.append("</row>");

            final int currentSheetIndex = sheetStack.peek();
            assertEquals(currentSheetIndex, sheetIndex);

            final int poppedRowNum = rowStack.pop();
            assertEquals(rowNum, poppedRowNum);
        }

        @Override
        public void onHandleCell(int sheetIndex, int rowNum, int columnNum, @NotNull CellValue cellValue) {
            // run for coverage
            WorkbookEventReader.EventHandler.super.onHandleCell(sheetIndex, rowNum, columnNum, cellValue);

            xml.append("<cell javaType=\"")
                    .append(cellValue.isNull() ? "null" : cellValue.originalType().getCanonicalName())
                    .append("\" index=\"")
                    .append(columnNum)
                    .append("\">")
                    .append(cellValue.isNull() ? "" : cellValue.originalValue())
                    .append("</cell>");

            final int currentSheetIndex = sheetStack.peek();
            assertEquals(currentSheetIndex, sheetIndex);

            final int currentRowNum = rowStack.peek();
            assertEquals(rowNum, currentRowNum);
        }
    }

}
