package com.github.kumasuke120.excel.util;

import com.github.kumasuke120.excel.WorkbookEventReader;
import org.junit.jupiter.api.Test;

import java.lang.invoke.MethodType;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.time.format.DateTimeFormatter;

import static org.junit.jupiter.api.Assertions.*;

class ReflectionUtilsTest {

    @Test
    void getClassTest() {
        assertNotNull(ReflectionUtils.getClass("java.lang.String"));
        assertNull(ReflectionUtils.getClass("com.github.kumasuke120.excel.util.NonExistentClass"));
    }

    @Test
    void findStaticGetterHandle() {
        assertNotNull(ReflectionUtils.findStaticGetterHandle(DateTimeFormatter.class, "ISO_LOCAL_DATE",
                DateTimeFormatter.class));
        assertNull(ReflectionUtils.findStaticGetterHandle(ReflectionUtils.class, "nonExistentField",
                String.class));
    }

    @Test
    void findStaticHandle() {
        assertNotNull(ReflectionUtils.findStaticHandle(ReflectionUtils.class, "getClass", MethodType.methodType(Class.class, String.class)));
        assertNull(ReflectionUtils.findStaticHandle(ReflectionUtils.class, "nonExistentMethod", MethodType.methodType(Class.class, String.class)));
    }

    @Test
    void findVirtualHandle() {
        assertNotNull(ReflectionUtils.findVirtualHandle(WorkbookEventReader.class, "read",
                MethodType.methodType(void.class, WorkbookEventReader.EventHandler.class)));
        assertNull(ReflectionUtils.findVirtualHandle(ReflectionUtils.class, "nonExistentMethod",
                MethodType.methodType(Class.class, String.class)));
    }

    @Test
    void newInstance() {
        try {
            final Constructor<ReflectionUtils> constructor = ReflectionUtils.class.getDeclaredConstructor();
            constructor.setAccessible(true);
            constructor.newInstance();
        } catch (InvocationTargetException e) {
            assertTrue(e.getTargetException() instanceof UnsupportedOperationException);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }
}