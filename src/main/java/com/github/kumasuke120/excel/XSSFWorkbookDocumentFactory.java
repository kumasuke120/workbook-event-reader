package com.github.kumasuke120.excel;

import com.github.kumasuke120.excel.util.ReflectionUtils;
import org.apache.xmlbeans.XmlException;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;

import java.io.IOException;
import java.io.InputStream;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodType;

/**
 * Factory class for creating and parsing `WorkbookDocument` instances.<br>
 * This class provides compatibility with multiple versions of Apache POI (starting from 3.17).
 */
@ApiStatus.Internal
class XSSFWorkbookDocumentFactory {

    private static final String FACTORY_CLASS_NAME = "org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument$Factory";
    private static final String FIELD_FACTORY_CLASS_NAME = "org.apache.xmlbeans.impl.schema.DocumentFactory";

    private static final Class<?> factoryClass;
    private static final Class<?> fieldFactoryClass;
    private static final MethodHandle fieldFactoryFieldHandle;
    private static final ParseFunction parseFunction;

    static {
        factoryClass = ReflectionUtils.getClass(FACTORY_CLASS_NAME);
        fieldFactoryClass = ReflectionUtils.getClass(FIELD_FACTORY_CLASS_NAME);
        fieldFactoryFieldHandle = getFieldFactoryFieldHandle();
        parseFunction = getParseFunction();
    }

    private XSSFWorkbookDocumentFactory() {
        throw new UnsupportedOperationException();
    }

    /**
     * Parses an InputStream into a `WorkbookDocument` instance.
     *
     * @param in the InputStream to parse
     * @return the parsed WorkbookDocument
     * @throws XmlException if an XML parsing error occurs
     * @throws IOException  if an I/O error occurs
     */
    @NotNull
    public static WorkbookDocument parse(@NotNull InputStream in) throws XmlException, IOException {
        if (parseFunction == null) {
            throw new UnsupportedOperationException("cannot find parse() method for WorkbookDocument.Factory");
        }
        return parseFunction.parse(in);
    }

    private static MethodHandle getFieldFactoryFieldHandle() {
        if (fieldFactoryClass == null) {
            return null;
        }
        return ReflectionUtils.findStaticGetterHandle(WorkbookDocument.class, "Factory", fieldFactoryClass);
    }

    private static MethodHandle getFactoryClassHandle() {
        return ReflectionUtils.findStaticHandle(factoryClass, "parse",
                MethodType.methodType(WorkbookDocument.class, InputStream.class));
    }

    private static MethodHandle getFieldFactoryHandle() {
        return ReflectionUtils.findVirtualHandle(fieldFactoryClass, "parse",
                MethodType.methodType(Object.class, InputStream.class));
    }

    private static ParseFunction getParseFunction() {
        if (factoryClass != null) {
            // for Apache POI 3.x and 4.x
            final MethodHandle factoryHandle = getFactoryClassHandle();
            if (factoryHandle == null) {
                return null;
            }
            return inStream -> invokeMethodHandle(factoryHandle, inStream);
        } else if (fieldFactoryFieldHandle != null) {
            // for Apache POI 5.x
            final MethodHandle fieldFactoryHandle = getFieldFactoryHandle();
            final Object factory = getFieldFactoryInstance();
            return inStream -> invokeMethodHandle(fieldFactoryHandle, factory, inStream);
        } else {
            return null;
        }
    }

    private static Object getFieldFactoryInstance() {
        Object factory;
        try {
            factory = fieldFactoryFieldHandle.invoke();
        } catch (Throwable e) {
            throw new AssertionError(e);
        }
        return factory;
    }

    @SuppressWarnings("unchecked")
    private static <T> T invokeMethodHandle(MethodHandle handle, Object... args) throws XmlException, IOException {
        try {
            return (T) handle.invokeWithArguments(args);
        } catch (Throwable e) {
            if (e instanceof XmlException) {
                throw (XmlException) e;
            } else if (e instanceof IOException) {
                throw (IOException) e;
            } else {
                throw new AssertionError(e);
            }
        }
    }

    /**
     * Functional interface for parsing an InputStream into a WorkbookDocument.
     */
    @FunctionalInterface
    private interface ParseFunction {

        /**
         * Parses an InputStream into a WorkbookDocument.
         *
         * @param in the InputStream to parse
         * @return the parsed WorkbookDocument
         * @throws XmlException if an XML parsing error occurs
         * @throws IOException  if an I/O error occurs
         */
        @NotNull
        WorkbookDocument parse(@NotNull InputStream in) throws XmlException, IOException;

    }

}
