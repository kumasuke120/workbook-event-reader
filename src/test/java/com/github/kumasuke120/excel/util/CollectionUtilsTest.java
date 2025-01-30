package com.github.kumasuke120.excel.util;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

@SuppressWarnings("ConstantValue")
class CollectionUtilsTest {

    @Test
    void newInstance() {
        try {
            final Constructor<CollectionUtils> constructor = CollectionUtils.class.getDeclaredConstructor();
            constructor.setAccessible(true);
            constructor.newInstance();
        } catch (InvocationTargetException e) {
            assertTrue(e.getTargetException() instanceof UnsupportedOperationException);
        } catch (ReflectiveOperationException e) {
            throw new AssertionError(e);
        }
    }

    @Test
    void isEmpty() {
        assertTrue(CollectionUtils.isEmpty(null));
        assertTrue(CollectionUtils.isEmpty(new java.util.ArrayList<>()));
        assertFalse(CollectionUtils.isEmpty(java.util.Arrays.asList(1, 2, 3)));
    }

    @Test
    void isNotEmpty() {
        assertFalse(CollectionUtils.isNotEmpty(null));
        assertFalse(CollectionUtils.isNotEmpty(new java.util.ArrayList<>()));
        assertTrue(CollectionUtils.isNotEmpty(java.util.Arrays.asList(1, 2, 3)));
    }

}