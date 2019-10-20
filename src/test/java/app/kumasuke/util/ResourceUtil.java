package app.kumasuke.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.nio.file.Paths;

public class ResourceUtil {
    private ResourceUtil() {
        throw new UnsupportedOperationException();
    }

    public static Path getPathOfClasspathResource(String resourceName) {
        final URL resourceURL = ClassLoader.getSystemResource(resourceName);
        if (resourceURL != null) {
            try {
                return Paths.get(ClassLoader.getSystemResource(resourceName).toURI());
            } catch (URISyntaxException e) {
                throw new AssertionError(e);
            }
        } else {
            throw new IllegalArgumentException();
        }
    }

    public static String loadClasspathResourceToString(String resourceName) throws IOException {
        final int bufferSize = 2048;

        final char[] buffer = new char[bufferSize];
        final StringBuilder builder = new StringBuilder();
        try (final InputStream is = ClassLoader.getSystemResourceAsStream(resourceName)) {
            assert is != null;

            try (final InputStreamReader isr = new InputStreamReader(is, StandardCharsets.UTF_8)) {
                int nChars;

                while ((nChars = isr.read(buffer)) != -1) {
                    builder.append(buffer, 0, nChars);
                }
            }
        }

        return builder.toString();
    }
}
