package app.kumasuke.excel;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class XSSFWorkbookEventReaderTest extends AbstractWorkbookEventReaderTest<XSSFWorkbookEventReader> {
    private static final String NORMAL_FILE_NAME = "workbook.xlsx";
    private static final String ENCRYPTED_FILE_NAME = "workbook-encrypted.xlsx";


    XSSFWorkbookEventReaderTest() {
        super(NORMAL_FILE_NAME, ENCRYPTED_FILE_NAME, XSSFWorkbookEventReader.class);
    }

    @Test
    @Override
    void read() {
        super.read();
    }

    @Test
    @Override
    void cancel() {
        super.cancel();

        dealWithReader(reader -> {
            assert reader instanceof XSSFWorkbookEventReader;

            final boolean[] cancelledRef = {false};
            final WorkbookEventReader.EventHandler handler = new WorkbookEventReader.EventHandler() {
                @Override
                public void onStartRow(int sheetIndex, int rowNum) {
                    reader.cancel();
                }

                @Override
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
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

    @Test
    @Override
    void close() {
        super.close();
    }

    @Test
    void setUse1904Windowing() {
        dealWithReader(reader -> {
            assert reader instanceof XSSFWorkbookEventReader;
            ((XSSFWorkbookEventReader) reader).setUse1904Windowing(true);

            reader.read(new WorkbookEventReader.EventHandler() {
                @Override
                public void onStartDocument() {
                    assertThrows(IllegalReaderStateException.class,
                                 () -> ((XSSFWorkbookEventReader) reader).setUse1904Windowing(false));
                }

                @Override
                public void onHandleCell(int sheetIndex, int rowNum, int columnNum, CellValue cellValue) {
                    if (sheetIndex == 0 && (rowNum == 3 || rowNum == 4) && columnNum == 1) {
                        if (!cellValue.isNull()) {
                            assertEquals(2022, cellValue.localDateValue().getYear());
                        }
                    }
                }
            });

            reader.close();

            assertThrows(IllegalReaderStateException.class,
                         () -> ((XSSFWorkbookEventReader) reader).setUse1904Windowing(false));
        });
    }
}
