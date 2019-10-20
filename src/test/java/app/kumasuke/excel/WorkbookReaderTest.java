package app.kumasuke.excel;

import app.kumasuke.util.ResourceUtil;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.ThreadLocalRandom;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

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
