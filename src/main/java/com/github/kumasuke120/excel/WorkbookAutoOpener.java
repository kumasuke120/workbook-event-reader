package com.github.kumasuke120.excel;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.util.IOUtils;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

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
@ApiStatus.Internal
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
    WorkbookAutoOpener(@NotNull Path filePath, @Nullable String password) {
        this(filePath, null, password);
    }

    /**
     * Creates a new {@link WorkbookAutoOpener} based on the given workbook {@link InputStream}
     * and the given password if possible.
     *
     * @param in       {@link InputStream} of the workbook to be read
     * @param password password to open the file
     */
    WorkbookAutoOpener(@NotNull InputStream in, @Nullable String password) {
        this(null, in, password);
    }

    private WorkbookAutoOpener(@Nullable Path filePath, @Nullable InputStream in, @Nullable String password) {
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
    @NotNull
    WorkbookEventReader open() {
        if (in == null && filePath != null) {
            return openByPath();
        } else if (in == null) {
            throw new NullPointerException();
        } else if (filePath == null) {
            return openByInputStream();
        } else {
            throw new NullPointerException();
        }
    }

    @NotNull
    private WorkbookEventReader openByInputStream() {
        assert in != null;

        try (InputStream stream = FileMagic.prepareToCheckMagic(in)) {
            final List<ReaderConstructor> constructors = buildConstructorsByMagic(stream);
            if (constructors.size() > 1) {
                return tryConstruct(constructors, createInSupplier(stream));
            } else {
                return tryConstruct(constructors, stream);
            }
        } catch (IOException e) {
            throw new WorkbookIOException("Cannot open and read workbook", e);
        }
    }

    private Supplier<InputStream> createInSupplier(InputStream stream) throws IOException {
        final byte[] bytes = IOUtils.toByteArray(stream);
        return () -> new ByteArrayInputStream(bytes);
    }

    @NotNull
    private List<ReaderConstructor> buildConstructorsByMagic(InputStream stream) throws IOException {
        final List<ReaderConstructor> constructors = new ArrayList<>();
        final FileMagic magic = FileMagic.valueOf(stream);
        switch (magic) {
            case OOXML:
                constructors.add(xssfConstructor);
                break;
            case OLE2:
                constructors.add(hssfConstructor);
                if (password != null) {
                    constructors.add(xssfConstructor);
                }
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

    @NotNull
    private WorkbookEventReader openByPath() {
        assert filePath != null;

        final String ext = getExtension();
        final List<ReaderConstructor> constructors = buildConstructorsByExtension(ext);
        return tryConstruct(constructors, filePath);
    }

    @NotNull
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

    @NotNull
    private List<ReaderConstructor> buildConstructorsByExtension(@NotNull String ext) {
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

    @NotNull
    private WorkbookEventReader tryConstruct(@NotNull List<ReaderConstructor> constructors,
                                             @NotNull InputStream stream) {
        return tryConstruct(constructors, () -> stream);
    }

    @NotNull
    private WorkbookEventReader tryConstruct(@NotNull List<ReaderConstructor> constructors,
                                             @NotNull Supplier<InputStream> inSupplier) {
        assert !constructors.isEmpty();

        WorkbookIOException thrown = null;
        for (ReaderConstructor constructor : constructors) {
            final InputStream in = inSupplier.get();
            try {
                return constructor.newInstance(in, password);
            } catch (WorkbookIOException e) {
                thrown = updateTryConstructException(thrown, e);
            }
        }

        // assert thrown != null;
        throw thrown;
    }

    @NotNull
    private WorkbookEventReader tryConstruct(@NotNull List<ReaderConstructor> constructors,
                                             @NotNull Path filePath) {
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

    @NotNull
    private WorkbookIOException updateTryConstructException(@Nullable WorkbookIOException thrown,
                                                            @NotNull WorkbookIOException e) {
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

    static class ReaderConstructor {
        private final Class<? extends WorkbookEventReader> readerClass;
        private final Constructor<? extends WorkbookEventReader> inputStreamConstructor;
        private final Constructor<? extends WorkbookEventReader> pathConstructor;

        ReaderConstructor(@NotNull Class<? extends WorkbookEventReader> readerClass) {
            this.readerClass = readerClass;
            this.inputStreamConstructor = findConstructor(InputStream.class);
            this.pathConstructor = findConstructor(Path.class);
        }

        @NotNull
        private Constructor<? extends WorkbookEventReader> findConstructor(@NotNull Class<?> firstParameterType) {
            try {
                return readerClass.getConstructor(firstParameterType, String.class);
            } catch (NoSuchMethodException e1) {
                try {
                    return readerClass.getConstructor(firstParameterType);
                } catch (NoSuchMethodException e2) {
                    throw new AssertionError("Shouldn't happen", e2);
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
        @NotNull
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
        @NotNull
        WorkbookEventReader newInstance(InputStream in, String password) {
            return newInstance(inputStreamConstructor, in, password);
        }

        @NotNull
        private WorkbookEventReader newInstance(@NotNull Constructor<? extends WorkbookEventReader> constructor,
                                                @NotNull Object firstParameter,
                                                @Nullable String password) {
            final int parameterCount = constructor.getParameterCount();
            final Object[] parameters = parameterCount == 1 ? new Object[]{firstParameter} :
                    new Object[]{firstParameter, password};
            try {
                return constructor.newInstance(parameters);
            } catch (InvocationTargetException e) {
                throw translateException(e);
            } catch (ReflectiveOperationException e) {
                throw new AssertionError("Shouldn't happen", e);
            }
        }

        @NotNull
        private WorkbookIOException translateException(@NotNull InvocationTargetException e) {
            final Throwable t = e.getTargetException();

            if (t instanceof Error) {
                throw (Error) t;
            }

            if (t instanceof WorkbookIOException) {
                return (WorkbookIOException) t;
            } else {
                return new WorkbookIOException("Cannot open workbook", t);
            }
        }
    }

}
