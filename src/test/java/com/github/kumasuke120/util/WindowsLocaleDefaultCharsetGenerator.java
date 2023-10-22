package com.github.kumasuke120.util;

import com.sun.jna.*;
import com.sun.jna.platform.win32.WTypes;
import com.sun.jna.platform.win32.WinDef;
import com.sun.jna.win32.W32APIOptions;

import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * @apiNote <a href="https://stackoverflow.com/questions/3864240/default-code-page-for-each-language-version-of-windows/7685385#7685385">stackoverflow.com</a>
 */
public class WindowsLocaleDefaultCharsetGenerator {

    private static final Map<String, Set<String>> additionalCharsetAliases;

    static {
        additionalCharsetAliases = new HashMap<>();
        register("windows-31j", "Shift_JIS");
        register("GBK", "GB18030");
        register("windows-874", "ISO-8859-11");
        register("x-windows-950", "Big5");
        register("windows-949", "EUC-KR");
    }

    /**
     * @apiNote <a href="https://learn.microsoft.com/en-us/windows/win32/api/winnls/nc-winnls-locale_enumprocex">learn.microsoft.com</a>
     */
    public static void main(String[] args) {
        final Map<String, Charset> localDefaultCharset = new HashMap<>();

        final Kernel32 kernel32 = Kernel32.INSTANCE;

        kernel32.EnumSystemLocalesEx((locale, dwFlags, lParam) -> {
            final int bufSize = 256;
            final char[] lpLCData = new char[bufSize];
            kernel32.GetLocaleInfoEx(locale, Kernel32.LOCALE_IDEFAULTANSICODEPAGE, lpLCData, bufSize);
            final WTypes.LPWSTR codePage = new WTypes.LPWSTR(new String(lpLCData));

            if (!codePage.getValue().equals("0")) {
                final Charset charset = findCharset(codePage.getValue());
                localDefaultCharset.put(locale.getValue(), charset);
            }

            return true;
        }, Kernel32.LOCALE_ALL, new WinDef.LPARAM(0), null);

        // outputs localeDefaultCharset codes
        System.out.println("final Map<String, Set<String>> ldc = new HashMap<>(" + localDefaultCharset.size() + ");");
        System.out.println();

        for (Map.Entry<String, Charset> e : localDefaultCharset.entrySet()) {
            final String locale = e.getKey();
            final Charset charset = e.getValue();

            Set<String> charsetNames = new HashSet<>(charset.aliases());
            charsetNames.add(charset.name());
            for (String charsetName : new HashSet<>(charsetNames)) {
                final Set<String> addition = additionalCharsetAliases.get(charsetName);
                if (addition != null) {
                    charsetNames.addAll(addition);
                }
            }

            final String params = charsetNames.stream().map(n -> "\"" + n + "\"").collect(Collectors.joining(", "));
            System.out.println("register(ldc, \"" + locale + "\", " + params + ");");
        }

        System.out.println();
        System.out.println("localeDefaultCharsets = Collections.unmodifiableMap(ldc);");
    }

    private static void register(String charset, String... aliases) {
        assert additionalCharsetAliases != null;

        for (String alias : aliases) {
            Charset.forName(alias);
            additionalCharsetAliases.computeIfAbsent(charset, k -> new HashSet<>()).add(alias);
        }
    }

    @SuppressWarnings("InjectedReferences")
    private static Charset findCharset(String codePage) {
        String charsetName = "windows-" + codePage;
        try {
            return Charset.forName(charsetName);
        } catch (IllegalArgumentException e) {
            charsetName = "cp" + codePage;
            return Charset.forName(charsetName);
        }
    }

    @SuppressWarnings({"SpellCheckingInspection", "UnusedReturnValue", "unused"})
    interface Kernel32 extends Library {

        Kernel32 INSTANCE = Native.load("kernel32", Kernel32.class, W32APIOptions.DEFAULT_OPTIONS);

        /**
         * enumerate all named based locales
         */
        int LOCALE_ALL = 0;

        /**
         * default ansi code page (use of Unicode is recommended instead)
         */
        int LOCALE_IDEFAULTANSICODEPAGE = 0x00001004;

        /**
         * @apiNote <a href="https://learn.microsoft.com/en-us/windows/win32/api/winnls/nc-winnls-locale_enumprocex">learn.microsoft.com</a>
         */
        interface LOCALE_ENUMPROCEX extends Callback {
            boolean invoke(WTypes.LPWSTR locale, int dwFlags, WinDef.LPARAM lParam);

        }

        /**
         * @apiNote <a href="https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-enumsystemlocalesex">learn.microsoft.com</a>
         */
        boolean EnumSystemLocalesEx(LOCALE_ENUMPROCEX lpLocaleEnumProcEx, int dwFlags, WinDef.LPARAM lParam, WinDef.LPVOID lpReserved);

        /**
         * @apiNote <a href="https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getlocaleinfoex">learn.microsoft.com</a>
         */
        int GetLocaleInfoEx(WTypes.LPWSTR lpLocaleName, int LCType, char[] lpLCData, int cchData);

    }

}
