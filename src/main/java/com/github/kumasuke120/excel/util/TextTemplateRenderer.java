package com.github.kumasuke120.excel.util;

import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;

import java.util.Map;

/**
 * Text template renderer, which can render strings based on string templates
 */
@ApiStatus.Internal
public class TextTemplateRenderer {

    private final String template;

    /**
     * Constructs a new {@code TextTemplateRenderer} with the specified string template.<br>
     *
     * @param template the string template
     */
    public TextTemplateRenderer(@NotNull String template) {
        this.template = template;
    }

    /**
     * Renders a string based on the specified string template and property set.<br>
     *
     * @param template the string template
     * @param props    the property set
     * @return the rendered string
     * @see #render(Map)
     */
    @NotNull
    public static String render(@NotNull String template, Map<String, ?> props) {
        return new TextTemplateRenderer(template).render(props);
    }

    /**
     * Renders a string based on the specified string template and property set.<br>
     *
     * <ol>
     * <li>Supports name replacement, e.g.: ({@code int value = 1}) <code>"a = ${value}"</code> will be rendered as <code>"a = 1"</code></li>
     * <li>Supports formatting, e.g.: ({@code double pi = 3.14159}) <code>"pi = ${pi:.2f}"</code> will be rendered as
     * <code>"pi = 3.14"</code>, refer to {@link String#format(String, Object...)} for specific formats</li>
     * <li>Special character escaping, i.e.: <code>"$$120"</code> will be rendered as <code>"$120"</code></li>
     * <li>During rendering, template errors do not throw exceptions and try to display as much as possible, e.g.: <code>"x = ${value"</code> will be rendered as <code>"x = ${value"</code></li>
     * <li>Non-existent property values will be displayed as empty strings, e.g.: () <code>"a = ${value}"</code> will be rendered as <code>"a = "</code></li>
     * </ol>
     *
     * @param props the property set
     * @return the rendered string
     */
    @NotNull
    public String render(@NotNull Map<String, ?> props) {
        final StringBuilder text = new StringBuilder();
        final StringBuilder propName = new StringBuilder();
        final StringBuilder propFormat = new StringBuilder();

        boolean doTextAppend;
        boolean doPropNameAppend;
        boolean doPropFormatAppend;

        ParseState s = ParseState.NORMAL;
        for (int i = 0; i < template.length(); i++) {
            final char c = template.charAt(i);

            doTextAppend = false;
            doPropNameAppend = false;
            doPropFormatAppend = false;
            switch (s) {
                case NORMAL: {
                    if ('$' == c) {
                        s = ParseState.MAYBE_PLACEHOLDER;
                    } else {
                        doTextAppend = true;
                    }
                    break;
                }
                case MAYBE_PLACEHOLDER: {
                    if ('{' == c) {
                        s = ParseState.IN_PLACEHOLDER;
                        propName.setLength(0);   // Initialize property name
                        propFormat.setLength(0); // Initialize property format
                    } else {
                        s = ParseState.NORMAL;
                        if ('$' != c) text.append("$");  // Support '$$' escaping to '$'
                        doTextAppend = true;
                    }
                    break;
                }
                case IN_PLACEHOLDER: {
                    if ('}' == c) {
                        s = ParseState.NORMAL;
                        text.append(getPropStringValue(props, propName, propFormat));
                    } else {
                        if (i == template.length() - 1) {   // Last character
                            text.append("${").append(propName).append(c);
                        } else if (':' == c) {
                            s = ParseState.MAYBE_PLACEHOLDER_FORMAT;
                        } else {
                            doPropNameAppend = true;
                        }
                    }
                    break;
                }
                case MAYBE_PLACEHOLDER_FORMAT: {
                    if ('}' == c) {
                        s = ParseState.IN_PLACEHOLDER; // Equivalent to no format, state rollback
                        i -= 1;                        // Character rollback
                    } else {
                        if (i == template.length() - 1) {   // Last character
                            text.append("${").append(propName).append(":").append(c);
                        } else {
                            s = ParseState.IN_PLACEHOLDER_FORMAT;
                            i -= 1; // Character rollback
                        }
                    }
                    break;
                }
                case IN_PLACEHOLDER_FORMAT: {
                    if ('}' == c) {
                        s = ParseState.IN_PLACEHOLDER; // Format reading ends, state rollback
                        i -= 1;                        // Character rollback
                    } else {
                        if (i == template.length() - 1) {
                            text.append("${").append(propName).append(":").append(propFormat).append(c);
                        } else {
                            doPropFormatAppend = true;
                        }
                    }
                    break;
                }
            }

            if (doTextAppend) text.append(c);
            if (doPropNameAppend) propName.append(c);
            if (doPropFormatAppend) propFormat.append(c);
        }

        return text.toString();
    }

    private String getPropStringValue(Map<String, ?> props, CharSequence propName, CharSequence format) {
        final Object prop = props.get(propName.toString());

        if (StringUtils.isBlank(format)) {
            format = "s";
        }

        return prop == null ? "" : String.format("%" + format, prop);
    }

    private enum ParseState {

        NORMAL,
        MAYBE_PLACEHOLDER,
        IN_PLACEHOLDER,
        MAYBE_PLACEHOLDER_FORMAT,
        IN_PLACEHOLDER_FORMAT

    }

}