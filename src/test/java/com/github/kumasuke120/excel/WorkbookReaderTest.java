package com.github.kumasuke120.excel;

import com.github.kumasuke120.util.ResourceUtil;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.ThreadLocalRandom;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.*;

@SuppressWarnings("EmptyTryBlock")
class WorkbookReaderTest {

    static final String CORRECT_PASSWORD = "password";

    static String randomWrongPassword() {
        String result;
        do {
            final StringBuilder builder = new StringBuilder(CORRECT_PASSWORD.length());
            for (int i = 0; i < CORRECT_PASSWORD.length(); i++) {
                final int pos = ThreadLocalRandom.current().nextInt(CORRECT_PASSWORD.length());
                builder.append(CORRECT_PASSWORD.charAt(pos));
            }
            result = builder.toString();
        } while (CORRECT_PASSWORD.equals(result));

        return result;
    }

    @SuppressWarnings("ConstantConditions")
    @Test
    void nullCheck() {
        assertThrows(NullPointerException.class, () -> {
            try (final WorkbookEventReader ignore = WorkbookEventReader.open((InputStream) null)) {
                // no-op
            }
        });
        assertThrows(NullPointerException.class, () -> {
            try (final WorkbookEventReader ignore = WorkbookEventReader.open((Path) null)) {
                // no-op
            }
        });
    }

    @Test
    void openWithPath() {
        final Path xlsPath = ResourceUtil.getPathOfClasspathResource("workbook.xls");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(xlsPath)) {
            assertTrue(reader instanceof HSSFWorkbookEventReader);
        }

        final Path xlsxPath = ResourceUtil.getPathOfClasspathResource("workbook.xlsx");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(xlsxPath)) {
            assertTrue(reader instanceof XSSFWorkbookEventReader);
        }

        final Path csvPath = ResourceUtil.getPathOfClasspathResource("ENGINES.csv");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(csvPath)) {
            assertTrue(reader instanceof CSVWorkbookEventReader);
        }

        final Path plainPath = ResourceUtil.getPathOfClasspathResource("workbook");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(plainPath)) {
            assertTrue(reader instanceof XSSFWorkbookEventReader);
        }

        assertThrows(WorkbookIOException.class, () -> {
            try (final WorkbookEventReader ignore = WorkbookEventReader.open(Paths.get("fileNotFound.xlsx"))) {
                // no-op
            }
        });
    }

    @Test
    void openWithPathAndPassword() {
        final Path xlsPath = ResourceUtil.getPathOfClasspathResource("workbook-encrypted.xls");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(xlsPath, CORRECT_PASSWORD)) {
            assertTrue(reader instanceof HSSFWorkbookEventReader);
        }
        assertThrows(WorkbookIOException.class, () -> {
            try (final WorkbookEventReader ignore = WorkbookEventReader.open(xlsPath)) {
                // no-op
            }
        });
        assertThrows(WorkbookIOException.class, () -> {
            try (final WorkbookEventReader ignore = WorkbookEventReader.open(xlsPath, randomWrongPassword())) {
                // no-op
            }
        });

        final Path xlsxPath = ResourceUtil.getPathOfClasspathResource("workbook-encrypted.xlsx");
        try (final WorkbookEventReader reader = WorkbookEventReader.open(xlsxPath, CORRECT_PASSWORD)) {
            assertTrue(reader instanceof XSSFWorkbookEventReader);
        }
        assertThrows(WorkbookIOException.class, () -> {
            try (final WorkbookEventReader ignore = WorkbookEventReader.open(xlsxPath)) {
                // no-op
            }
        });
        assertThrows(WorkbookIOException.class, () -> {
            try (final WorkbookEventReader ignore = WorkbookEventReader.open(xlsxPath, randomWrongPassword())) {
                // no-op
            }
        });
    }

    @SuppressWarnings("ConstantConditions")
    @Test
    void openWithStream() throws IOException {
        try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook.xls")) {
            try (final WorkbookEventReader reader = WorkbookEventReader.open(in)) {
                assertTrue(reader instanceof HSSFWorkbookEventReader);
            }
        }

        try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook.xlsx")) {
            try (final WorkbookEventReader reader = WorkbookEventReader.open(in)) {
                assertTrue(reader instanceof XSSFWorkbookEventReader);
            }
        }

        try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook.xlsx")) {
            final InputStream mockIn = markNotSupportedInputStream(in);
            try (final WorkbookEventReader reader = WorkbookEventReader.open(mockIn)) {
                assertTrue(reader instanceof XSSFWorkbookEventReader);
            }
        }

        try (final InputStream in = ClassLoader.getSystemResourceAsStream("ENGINES.csv")) {
            try (final WorkbookEventReader reader = WorkbookEventReader.open(in)) {
                assertTrue(reader instanceof CSVWorkbookEventReader);
            }
        }

        try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook")) {
            try (final WorkbookEventReader reader = WorkbookEventReader.open(in)) {
                assertTrue(reader instanceof XSSFWorkbookEventReader);
            }
        }

        try (final InputStream in = ClassLoader.getSystemResourceAsStream("sample-output.xml")) {
            assertThrows(WorkbookIOException.class, () -> {
                try (final WorkbookEventReader ignore = WorkbookEventReader.open(in)) {
                    // no-op
                }
            });
        }
    }

    private InputStream markNotSupportedInputStream(InputStream in) throws IOException {
        final InputStream mockIn = spy(in);
        when(mockIn.markSupported()).thenReturn(false);
        doNothing().when(mockIn).mark(anyInt());
        doNothing().when(mockIn).reset();
        return mockIn;
    }

    @SuppressWarnings("ConstantConditions")
    @Test
    void openWithStreamAndPassword() throws IOException {
        try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook-encrypted.xls")) {
            try (final WorkbookEventReader reader = WorkbookEventReader.open(in, CORRECT_PASSWORD)) {
                assertTrue(reader instanceof HSSFWorkbookEventReader);
            }
        }
        assertThrows(WorkbookIOException.class, () -> {
            try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook-encrypted.xls")) {
                try (final WorkbookEventReader ignore = WorkbookEventReader.open(in)) {
                    // no-op
                }
            }
        });
        assertThrows(WorkbookIOException.class, () -> {
            try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook-encrypted.xls")) {
                try (final WorkbookEventReader ignore = WorkbookEventReader.open(in, randomWrongPassword())) {
                    // no-op
                }
            }
        });

        try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook-encrypted.xlsx")) {
            try (final WorkbookEventReader reader = WorkbookEventReader.open(in, CORRECT_PASSWORD)) {
                assertTrue(reader instanceof XSSFWorkbookEventReader);
            }
        }
        assertThrows(WorkbookIOException.class, () -> {
            try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook-encrypted.xlsx")) {
                try (final WorkbookEventReader ignore = WorkbookEventReader.open(in)) {
                    // no-op
                }
            }
        });
        assertThrows(WorkbookIOException.class, () -> {
            try (final InputStream in = ClassLoader.getSystemResourceAsStream("workbook-encrypted.xlsx")) {
                try (final WorkbookEventReader ignore = WorkbookEventReader.open(in, randomWrongPassword())) {
                    // no-op
                }
            }
        });
    }

}
