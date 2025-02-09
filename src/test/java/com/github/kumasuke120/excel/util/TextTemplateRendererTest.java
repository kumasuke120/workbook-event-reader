package com.github.kumasuke120.excel.util;

import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

class TextTemplateRendererTest {

    @Test
    void testRender() {
        final String template = "a = ${value}";
        final Map<String, Object> props = Collections.singletonMap("value", 1);
        final String rendered = TextTemplateRenderer.render(template, props);
        assertEquals("a = 1", rendered);
    }

    @Test
    void testRenderWithFormat() {

        final String template = "pi = ${pi:.2f}";
        final Map<String, Object> props = Collections.singletonMap("pi", 3.14159);
        final String rendered = TextTemplateRenderer.render(template, props);
        assertEquals("pi = 3.14", rendered);

        final String template2 = "pi = ${pi:}";
        final String rendered2 = TextTemplateRenderer.render(template2, props);
        assertEquals("pi = 3.14159", rendered2);

        final String template3 = "pi = ${pi:.2f";
        final String rendered3 = TextTemplateRenderer.render(template3, props);
        assertEquals("pi = ${pi:.2f", rendered3);

        final String template4 = "pi = ${pi:d";
        final String rendered4 = TextTemplateRenderer.render(template4, props);
        assertEquals("pi = ${pi:d", rendered4);
    }

    @Test
    void testRenderWithEscape() {
        final String template = "$120";
        final Map<String, Object> props = new HashMap<>(0);
        final String rendered = TextTemplateRenderer.render(template, props);
        assertEquals("$120", rendered);

        final String template2 = "$$120";
        final String rendered2 = TextTemplateRenderer.render(template2, props);
        assertEquals("$120", rendered2);
    }

    @Test
    void testRenderWithTemplateError() {
        final String template = "x = ${value";
        final Map<String, Object> props = new HashMap<>(0);
        final String rendered = TextTemplateRenderer.render(template, props);
        assertEquals("x = ${value", rendered);
    }

    @Test
    void testRenderWithNonExistentProperty() {
        final String template = "a = ${value}";
        final Map<String, Object> props = new HashMap<>(0);
        final String rendered = TextTemplateRenderer.render(template, props);
        assertEquals("a = ", rendered);
    }

}