package com.github.kumasuke120.excel;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.util.IOUtils;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Supplier;

/**
 * A {@link WorkbookAutoOpener} which creates a {@link WorkbookEventReader} based on the given parameters
 */
class WorkbookAutoOpener {

    private static final ReaderConstructor hssfConstructor = new ReaderConstructor(HSSFWorkbookEventReader.class);
    private static final ReaderConstructor xssfConstructor = new ReaderConstructor(XSSFWorkbookEventReader.class);
    private static final ReaderConstructor csvConstructor = new ReaderConstructor(CSVWorkbookEventReader.class);

    private final Path filePath;
    private final InputStream in;
    private final String password;

    /**
     * Creates a new {@link WorkbookAutoOpener} based on the given file path using given password.
     *
     * @param filePath {@link Path} of the workbook to be opened
     * @param password password to open the file
     */
    WorkbookAutoOpener(Path filePath, String password) {
        this(filePath, null, password);
    }

    /**
     * Creates a new {@link WorkbookAutoOpener} based on thegiven workbook {@link InputStream}
     * and the given password if possible.
     *
     * @param in       {@link InputStream} of the workbook to be read
     * @param password password to open the file
     */
    WorkbookAutoOpener(InputStream in, String password) {
        this(null, in, password);
    }

    private WorkbookAutoOpener(Path filePath, InputStream in, String password) {
        this.in = in;
        this.filePath = filePath;
        this.password = password;
    }

    /**
     * Opens the specified file with an appropriate {@link WorkbookEventReader} if possible.
     *
     * @return {@link WorkbookEventReader} to read the specified file
     * @throws NullPointerException one of the required parameters is <code>null</code>
     * @throws WorkbookIOException  errors happened when opening
     */
    WorkbookEventReader open() {
        if (in == null && filePath != null) {
            return openByPath();
        } else if (in == null) {
            throw new NullPointerException();
        } else if (filePath == null) {
            return openByInputStream();
        } else {
            throw new AssertionError("This shouldn't happen");
        }
    }

    private WorkbookEventReader openByInputStream() {
        assert in != null;

        final byte[] bytes;
        try {
            bytes = IOUtils.toByteArray(in);
        } catch (IOException e) {
            throw new WorkbookIOException("Cannot open and read workbook", e);
        }

        final List<ReaderConstructor> constructors = buildConstructorsByMagic(bytes);
        final Supplier<InputStream> inSupplier = () -> new ByteArrayInputStream(bytes);
        return tryConstruct(constructors, inSupplier);
    }

    private FileMagic checkFileMagic(byte[] bytes) {
        try (final ByteArrayInputStream in = new ByteArrayInputStream(bytes);
             final InputStream stream = FileMagic.prepareToCheckMagic(in)) {
            return FileMagic.valueOf(stream);
        } catch (IOException e) {
            return FileMagic.UNKNOWN;
        }
    }

    private List<ReaderConstructor> buildConstructorsByMagic(byte[] bytes) {
        final List<ReaderConstructor> constructors = new ArrayList<>();
        final FileMagic magic = checkFileMagic(bytes);
        switch (magic) {
            case OOXML:
                constructors.add(xssfConstructor);
                break;
            case UNKNOWN:
                constructors.add(csvConstructor);
                // fallthrough
            default:
                constructors.add(xssfConstructor);
                constructors.add(hssfConstructor);
                break;
        }
        return constructors;
    }

    private WorkbookEventReader openByPath() {
        assert filePath != null;

        final String ext = getExtension();
        final List<ReaderConstructor> constructors = buildConstructorsByExtension(ext);
        return tryConstruct(constructors, filePath);
    }

    private String getExtension() {
        assert filePath != null;

        final Path fileName = filePath.getFileName();
        if (fileName == null) {
            return "";
        }

        final String fileNameStr = fileName.toString();
        final int extPos = fileNameStr.lastIndexOf(".");
        return extPos == -1 ? "" : fileNameStr.substring(extPos + 1);
    }

    private List<ReaderConstructor> buildConstructorsByExtension(String ext) {
        final List<ReaderConstructor> constructors = new ArrayList<>();
        switch (ext) {
            case "xlsx":
                constructors.add(xssfConstructor);
                constructors.add(hssfConstructor);
                break;
            case "xls":
                constructors.add(hssfConstructor);
                constructors.add(xssfConstructor);
                break;
            case "csv":
                constructors.add(csvConstructor);
                break;
            default:
                constructors.add(xssfConstructor);
                constructors.add(hssfConstructor);
                constructors.add(csvConstructor);
                break;
        }
        return constructors;
    }

    private WorkbookEventReader tryConstruct(List<ReaderConstructor> constructors, Supplier<InputStream> inSupplier) {
        assert !constructors.isEmpty();

        WorkbookIOException thrown = null;
        for (ReaderConstructor constructor : constructors) {
            final InputStream in = inSupplier.get();
            try {
                return constructor.newInstance(in, password);
            } catch (WorkbookIOException e) {
                // EncryptedDocumentException means password incorrect, it should be thrown directly
                thrown = updateTryConstructException(thrown, e);
            }
        }

        // assert thrown != null;
        throw thrown;
    }

    private WorkbookEventReader tryConstruct(List<ReaderConstructor> constructors, Path filePath) {
        assert !constructors.isEmpty();

        WorkbookIOException thrown = null;
        for (ReaderConstructor constructor : constructors) {
            try {
                return constructor.newInstance(filePath, password);
            } catch (WorkbookIOException e) {
                thrown = updateTryConstructException(thrown, e);
            }
        }

        // assert thrown != null;
        throw thrown;
    }

    private WorkbookIOException updateTryConstructException(WorkbookIOException thrown, WorkbookIOException e) {
        // EncryptedDocumentException means password incorrect, it should be thrown directly
        if (e.getCause() instanceof EncryptedDocumentException) {
            throw e;
        } else {
            if (thrown != null) {
                e.addSuppressed(thrown);
            }
            thrown = e;
        }
        return thrown;
    }

    private static class ReaderConstructor {
        private final Class<? extends WorkbookEventReader> readerClass;
        private final Constructor<? extends WorkbookEventReader> inputStreamConstructor;
        private final Constructor<? extends WorkbookEventReader> pathConstructor;

        ReaderConstructor(Class<? extends WorkbookEventReader> readerClass) {
            this.readerClass = readerClass;
            this.inputStreamConstructor = findConstructor(InputStream.class);
            this.pathConstructor = findConstructor(Path.class);
        }

        private Constructor<? extends WorkbookEventReader> findConstructor(Class<?> firstParameterType) {
            try {
                return readerClass.getConstructor(firstParameterType, String.class);
            } catch (NoSuchMethodException e1) {
                try {
                    return readerClass.getConstructor(firstParameterType);
                } catch (NoSuchMethodException e2) {
                    throw new AssertionError("This shouldn't happen", e2);
                }
            }
        }

        /**
         * Creates a new {@link WorkbookEventReader} based on the given file path using given password.
         *
         * @param filePath file path of the workbook
         * @param password password to open the file
         * @throws NullPointerException <code>filePath</code> is <code>null</code>
         * @throws WorkbookIOException  errors happened when opening
         */
        WorkbookEventReader newInstance(Path filePath, String password) {
            return newInstance(pathConstructor, filePath, password);
        }

        /**
         * Creates a new {@link WorkbookEventReader} based on the given workbook {@link InputStream}
         * and the given password if possible.
         *
         * @param in       {@link InputStream} of the workbook to be read
         * @param password password to open the file
         * @throws NullPointerException <code>filePath</code> is <code>null</code>
         * @throws WorkbookIOException  errors happened when opening
         */
        WorkbookEventReader newInstance(InputStream in, String password) {
            return newInstance(inputStreamConstructor, in, password);
        }

        private WorkbookEventReader newInstance(Constructor<? extends WorkbookEventReader> constructor,
                                                Object firstParameter, String password) {
            final int parameterCount = constructor.getParameterCount();
            final Object[] parameters = parameterCount == 1 ? new Object[]{firstParameter} :
                    new Object[]{firstParameter, password};
            try {
                return constructor.newInstance(parameters);
            } catch (InvocationTargetException e) {
                throw translateException(e);
            } catch (ReflectiveOperationException e) {
                throw new AssertionError("This shouldn't happen", e);
            }
        }

        private WorkbookIOException translateException(InvocationTargetException e) {
            final Throwable t = e.getTargetException();
            if (t instanceof WorkbookIOException) {
                return (WorkbookIOException) t;
            } else {
                return new WorkbookIOException(t);
            }
        }
    }

}
