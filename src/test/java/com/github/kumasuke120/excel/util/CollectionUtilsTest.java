package com.github.kumasuke120.excel.util;

import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

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
        // test for collection
        assertTrue(CollectionUtils.isEmpty((Collection<?>) null));
        assertTrue(CollectionUtils.isEmpty(new java.util.ArrayList<>()));
        assertFalse(CollectionUtils.isEmpty(java.util.Arrays.asList(1, 2, 3)));

        // test for map
        assertTrue(CollectionUtils.isEmpty((java.util.Map<?, ?>) null));
        assertTrue(CollectionUtils.isEmpty(new java.util.HashMap<>()));
        Map<String, String> map = new HashMap<>();
        map.put("1", "one");
        map.put("2", "two");
        assertFalse(CollectionUtils.isEmpty(map));
    }

    @Test
    void isNotEmpty() {
        // test for collection
        assertFalse(CollectionUtils.isNotEmpty((Collection<?>) null));
        assertFalse(CollectionUtils.isNotEmpty(new java.util.ArrayList<>()));
        assertTrue(CollectionUtils.isNotEmpty(java.util.Arrays.asList(1, 2, 3)));

        // test for map
        assertFalse(CollectionUtils.isNotEmpty((java.util.Map<?, ?>) null));
        assertFalse(CollectionUtils.isNotEmpty(new java.util.HashMap<>()));
        Map<String, String> map = new HashMap<>();
        map.put("1", "one");
        map.put("2", "two");
        assertTrue(CollectionUtils.isNotEmpty(map));
    }

}