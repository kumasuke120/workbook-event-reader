package com.github.kumasuke120.excel.util;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

@SuppressWarnings("ConstantValue")
class CollectionUtilsTest {

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