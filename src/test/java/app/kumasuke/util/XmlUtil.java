package app.kumasuke.util;

import nu.xom.*;

import java.io.IOException;
import java.util.*;

public class XmlUtil {
    private XmlUtil() {
        throw new UnsupportedOperationException();
    }

    public static boolean isSameXml(String aXml, String bXml) {
        if (Objects.equals(aXml, bXml)) {
            return true;
        } else {
            final Document aDoc = buildSilently(aXml);
            final Document bDoc = buildSilently(bXml);

            final Element aRoot = aDoc.getRootElement();
            final Element bRoot = bDoc.getRootElement();

            return isSameElement(aRoot, bRoot);
        }
    }

    private static boolean isSameElement(Element aElement, Element bElement) {
        final String aQN = aElement.getQualifiedName();
        final String bQN = bElement.getQualifiedName();

        if (!Objects.equals(aQN, bQN)) {
            System.err.println("Element qualified name: " + aQN + " != " + bQN);
            return false;
        } else {
            final int aAttrCount = aElement.getAttributeCount();
            final int bAttrCount = bElement.getAttributeCount();

            if (aAttrCount != bAttrCount) {
                System.err.println("Element attribute count <" + aQN + "> : " + aAttrCount + " != " + bAttrCount);
                return false;
            } else {
                final Map<String, String> aAttrs = extractAttributes(aElement);
                final Map<String, String> bAttrs = extractAttributes(bElement);

                if (!Objects.equals(aAttrs, bAttrs)) {
                    System.err.println("Element attribute(s) <" + aQN + "> : " + aAttrs + " != " + bAttrs);
                    return false;
                }

                final int aChildCount = aElement.getChildCount();
                final int bChildCount = aElement.getChildCount();

                if (aChildCount != bChildCount) {
                    System.err.println("Element child count <" + aQN + "> : " + aChildCount + " != " + bChildCount);
                    return false;
                } else if (aChildCount > 0) {
                    if (aChildCount == 1) {
                        final Node aChild = aElement.getChild(0);
                        final Node bChild = bElement.getChild(0);

                        if (aChild instanceof Text && bChild instanceof Text) {
                            final String aValue = aElement.getValue();
                            final String bValue = bElement.getValue();
                            final boolean valueEquals = Objects.equals(aValue, bValue);
                            if (!valueEquals) {
                                System.err.println("Element child value <" + aQN + "> : " + aValue + " != " + bValue);
                            }
                            return valueEquals;
                        }
                    }

                    final Map<String, List<Element>> aChildElements = extractChildElements(aElement);
                    final Map<String, List<Element>> bChildElements = extractChildElements(bElement);

                    if (aChildElements.size() != bChildElements.size()) {
                        System.err.println("Element child count <" + aQN + "> : " +
                                                   aChildElements.size() + " != " + bChildElements.size());
                        return false;
                    }

                    for (final String name : aChildElements.keySet()) {
                        final List<Element> aElements = aChildElements.get(name);
                        final List<Element> bElements = bChildElements.get(name);

                        final int aElementsSize = aElements.size();
                        final int bElementsSize = bElements == null ? 0 : bElements.size();
                        if (aElementsSize != bElementsSize) {
                            System.err.println("Child element count <" + name + "> : " +
                                                       aElementsSize + " != " + bElementsSize +
                                                       System.lineSeparator() +
                                                       "=".repeat(80) + System.lineSeparator() +
                                                       aElement.toXML() + System.lineSeparator() +
                                                       "-".repeat(80) + System.lineSeparator() +
                                                       bElement.toXML() +
                                                       System.lineSeparator() +
                                                       "=".repeat(80));
                            return false;
                        }

                        for (int i = 0; i < aElements.size(); i++) {
                            final Element aChildElement = aElements.get(i);
                            final Element bChildElement = bElements.get(i);

                            if (!isSameElement(aChildElement, bChildElement)) {
                                return false;
                            }
                        }
                    }
                }

                return true;
            }
        }
    }

    private static Map<String, String> extractAttributes(Element element) {
        final Map<String, String> result = new HashMap<>();

        for (int i = 0; i < element.getAttributeCount(); i++) {
            final Attribute attr = element.getAttribute(i);
            result.put(attr.getQualifiedName(), attr.getValue());
        }

        return result;
    }

    private static Map<String, List<Element>> extractChildElements(Element element) {
        final Map<String, List<Element>> result = new HashMap<>();

        final Elements childElements = element.getChildElements();
        for (int i = 0; i < childElements.size(); i++) {
            final Element e = childElements.get(i);
            final String name = e.getQualifiedName();
            final List<Element> elements = result.computeIfAbsent(name, k -> new ArrayList<>());
            elements.add(e);
        }

        return result;
    }

    private static Document buildSilently(String xml) {
        final var builder = new Builder();
        try {
            return builder.build(xml, null);
        } catch (ParsingException | IOException e) {
            throw new AssertionError(e);
        }
    }
}
