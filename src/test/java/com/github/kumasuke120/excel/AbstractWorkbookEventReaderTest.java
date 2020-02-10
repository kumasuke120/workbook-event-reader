package com.github.kumasuke120.excel;

import com.github.kumasuke120.util.ResourceUtil;
import com.github.kumasuke120.util.XmlUtil;
import org.apache.poi.EncryptedDocumentException;
import org.junit.jupiter.api.Assertions;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Path;
import java.util.Stack;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

abstract class AbstractWorkbookEventReaderTest<R extends AbstractWorkbookEventReader> {

    private final String normalFileName;
    private final String encryptedFileName;
    private final Class<R> readerClass;

    AbstractWorkbookEventReaderTest(String normalFileName,
                                    String encryptedFileName,
                                    Class<R> readerClass) {
        this.normalFileName = normalFileName;
        this.encryptedFileName = encryptedFileName;
        this.readerClass = readerClass;
    }

    @SuppressWarnings("EmptyTryBlock")
    final void dealWithReader(Consumer<WorkbookEventReader> consumer) {
        // constructor(Path)
        final Path filePath = ResourceUtil.getPathOfClasspathResource(normalFileName);
        try (final WorkbookEventReader reader = pathConstructor().newInstance(filePath)) {
            consumer.accept(reader);
        }
        assertThrows(NullPointerException.class, () -> {
            try (final WorkbookEventReader ignore = pathConstructor().newInstance((Object) null)) {
                // no-op
            }
        });

        // constructor(InputStream)
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

        // constructor(Path, String)
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

        // constructor(InputStream, String)
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
    }

    private LightWeightConstructor<R> pathConstructor() {
        try {
            return new LightWeightConstructor<>(readerClass.getConstructor(Path.class));
        } catch (NoSuchMethodException e) {
            throw new AssertionError(e);
        }
    }

    private LightWeightConstructor<R> pathAndPasswordConstructor() {
        try {
            return new LightWeightConstructor<>(readerClass.getConstructor(Path.class, String.class));
        } catch (NoSuchMethodException e) {
            throw new AssertionError(e);
        }
    }

    private LightWeightConstructor<R> inputStreamConstructor() {
        try {
            return new LightWeightConstructor<>(readerClass.getConstructor(InputStream.class));
        } catch (NoSuchMethodException e) {
            throw new AssertionError(e);
        }
    }

    private LightWeightConstructor<R> inputStreamAndPasswordConstructor() {
        try {
            return new LightWeightConstructor<>(readerClass.getConstructor(InputStream.class, String.class));
        } catch (NoSuchMethodException e) {
            throw new AssertionError(e);
        }
    }

    void read() {
        dealWithReader(reader -> {
            final TestEventHandler handler = new TestEventHandler();
            reader.read(handler);

            final String xml = handler.getXml();
            assertSameWithSample("sample-output.xml", xml);

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
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
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
            assertSameWithSample("sample-output-2.xml", xml);

            // reader has been cancelled
            assertTrue(cancelledRef[0]);

            // reader is not being read
            assertThrows(IllegalReaderStateException.class, reader::cancel);
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

    private static class LightWeightConstructor<T> {
        private final Constructor<T> constructor;

        LightWeightConstructor(Constructor<T> constructor) {
            this.constructor = constructor;
        }

        T newInstance(Object... initArgs) {
            try {
                return constructor.newInstance(initArgs);
            } catch (InvocationTargetException e) {
                final Throwable targetException = e.getTargetException();
                if (targetException instanceof RuntimeException) {
                    throw (RuntimeException) targetException;
                } else {
                    throw new AssertionError(e);
                }
            } catch (ReflectiveOperationException e) {
                throw new AssertionError(e);
            }
        }
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
        public void onStartSheet(int sheetIndex, String sheetName) {
            xml.append("<sheet index=\"")
                    .append(sheetIndex)
                    .append("\" name=\"")
                    .append(sheetName)
                    .append("\">");

            sheetStack.push(sheetIndex);

            if (sheetIndex == 0) {
                assertEquals("Sheet1", sheetName);
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
        public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
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
