package com.github.kumasuke120.excel;

import com.ibm.icu.text.CharsetDetector;
import com.ibm.icu.text.CharsetMatch;
import org.jetbrains.annotations.ApiStatus;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;
import org.mozilla.universalchardet.UniversalDetector;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.*;
import java.util.stream.Collectors;

/**
 * A character set detector that reads bytes to determine the fittest <code>Charset</code> for decoding these bytes
 */
@ApiStatus.Internal
class ByteArrayCharsetDetector {

    private final byte[] input;

    ByteArrayCharsetDetector(byte[] input) {
        this.input = input;
    }

    /**
     * Detects the fittest <code>Charset</code> for decoding the given bytes
     *
     * @return the detected <code>Charset</code>
     */
    @Nullable
    Charset detect() {
        final List<Map<String, CharsetCandidate>> candidatesList = new ArrayList<>();
        candidatesList.add(detectByICU());
        candidatesList.add(detectByLocale());
        candidatesList.add(detectByUniversal());

        final List<CharsetCandidate> charsetCandidates = mergeCandidates(candidatesList);
        return getWithMaxScore(charsetCandidates);
    }

    private Charset getWithMaxScore(@NotNull List<CharsetCandidate> candidates) {
        if (candidates.isEmpty()) {
            return null;
        }

        List<CharsetCandidate> orderedCandidates = new ArrayList<>(candidates);
        orderedCandidates.sort(Comparator.comparingInt(cc -> -cc.score));

        for (CharsetCandidate candidate : orderedCandidates) {
            try {
                return Charset.forName(candidate.charsetName);
            } catch (IllegalArgumentException ignored) {
            }
        }

        return null;
    }

    @NotNull
    private List<CharsetCandidate> mergeCandidates(@NotNull List<Map<String, CharsetCandidate>> candidatesList) {
        if (candidatesList.isEmpty()) {
            return new ArrayList<>(0);
        }

        Map<String, CharsetCandidate> result = new HashMap<>();

        for (Map<String, CharsetCandidate> candidates : candidatesList) {
            for (Map.Entry<String, CharsetCandidate> e : candidates.entrySet()) {
                final String charsetName = e.getKey();
                final CharsetCandidate candidate = e.getValue();

                final CharsetCandidate oldCandidate = result.get(charsetName);
                if (oldCandidate == null) {
                    result.put(charsetName, candidate);
                } else {
                    oldCandidate.score += candidate.score;
                    oldCandidate.sources.addAll(candidate.sources);
                }
            }
        }

        final ArrayList<CharsetCandidate> retCandidates = new ArrayList<>(result.values());
        for (CharsetCandidate candidate : retCandidates) {
            candidate.score *= candidate.sources.size();
        }

        return retCandidates;
    }

    private Map<String, CharsetCandidate> detectByUniversal() {
        try (final ByteArrayInputStream in = new ByteArrayInputStream(input)) {
            final String charset = UniversalDetector.detectCharset(in);
            if (charset == null) {
                return new HashMap<>(0);
            }

            final CharsetCandidate cc = new CharsetCandidate();

            cc.charsetName = charset;
            cc.score = 10;

            List<String> sources = new ArrayList<>();
            sources.add(UniversalDetector.class.getName());
            cc.sources = sources;

            final Map<String, CharsetCandidate> candidates = new HashMap<>(1);
            candidates.put(charset, cc);
            return candidates;
        } catch (IOException e) {
            return new HashMap<>(0);
        }
    }

    private Map<String, CharsetCandidate> detectByLocale() {
        final Locale locale = Locale.getDefault();

        final Set<String> codePages = WindowsCodePage.getByLocale(locale);
        if (codePages.isEmpty()) {
            return new HashMap<>(0);
        }

        final Map<String, CharsetCandidate> candidates = new HashMap<>(codePages.size());

        for (String cp : codePages) {
            final CharsetCandidate cc = new CharsetCandidate();
            cc.charsetName = cp;
            cc.score = 10;

            List<String> sources = new ArrayList<>();
            sources.add(Locale.class.getName());
            cc.sources = sources;

            candidates.put(cp, cc);
        }

        return candidates;
    }

    private Map<String, CharsetCandidate> detectByICU() {
        CharsetDetector cd = new CharsetDetector();
        cd.setText(input);

        final CharsetMatch[] matches = cd.detectAll();
        if (matches.length == 0) {
            return new HashMap<>(0);
        }

        // counts the occurrences by the language of the charset
        final Map<String, Long> countByLang = Arrays.stream(matches)
                .filter(cm -> cm.getLanguage() != null)
                .collect(Collectors.groupingBy(CharsetMatch::getLanguage, Collectors.counting()));

        final Map<String, CharsetCandidate> candidates = new HashMap<>(matches.length);

        for (final CharsetMatch match : matches) {
            final CharsetCandidate cc = new CharsetCandidate();

            cc.charsetName = match.getName();

            // calculates and sets score of each match
            int baseScore = match.getConfidence();
            final String lang = match.getLanguage();
            final long langCount = countByLang.getOrDefault(lang, 0L);
            baseScore += (int) ((langCount - 1) * 5);
            baseScore = Math.min(100, baseScore);
            cc.score = baseScore;

            List<String> sources = new ArrayList<>();
            sources.add(CharsetDetector.class.getName());
            cc.sources = sources;

            candidates.put(cc.charsetName, cc);
        }

        return candidates;
    }

    private static class WindowsCodePage {

        private static final Map<String, Set<String>> localeDefaultCharsets;

        static {
            final Map<String, Set<String>> ldc = new HashMap<>(521);

            register(ldc, "", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ky-KG", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "sw-TZ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-TD", "windows-1256", "cp1256");
            register(ldc, "es-BO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fil", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-SY", "windows-1256", "cp1256");
            register(ldc, "iu-Latn-CA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sr-Latn-BA", "windows-1250", "cp1250", "cp5346");
            register(ldc, "sw-UG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zu-ZA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-TN", "windows-1256", "cp1256");
            register(ldc, "hu-HU", "windows-1250", "cp1250", "cp5346");
            register(ldc, "af", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-RW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-CR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ru-UA", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "es-CL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-CO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar", "windows-1256", "cp1256");
            register(ldc, "en-SB", "windows-1252", "cp1252", "cp5348");
            register(ldc, "pt-PT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-SD", "windows-1252", "cp1252", "cp5348");
            register(ldc, "bs-Latn", "windows-1250", "cp1250", "cp5346");
            register(ldc, "en-SC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-CU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "pap-029", "windows-1252", "cp1252", "cp5348");
            register(ldc, "cy-GB", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-SH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "az", "cp5350", "windows-1254", "cp1254");
            register(ldc, "en-SG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-SL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ba", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "be", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "bg", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "zh-MO_stroke", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "pt-MZ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "af-ZA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ku-Arab-IQ", "windows-1256", "cp1256");
            register(ldc, "br", "windows-1252", "cp1252", "cp5348");
            register(ldc, "bs", "windows-1250", "cp1250", "cp5346");
            register(ldc, "fr-GF", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-CN_phoneb", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "pt-MO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-GN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-PG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-GQ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-GP", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-PH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ca", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-PK", "windows-1252", "cp1252", "cp5348");
            register(ldc, "az-Cyrl", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "sr-Cyrl-XK", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "en-PN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ko-KR", "x-windows-949", "ms_949", "windows949", "EUC-KR", "ms949", "windows-949");
            register(ldc, "en-PR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-TW_pronun", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "en-PW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "co-FR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "co", "windows-1252", "cp1252", "cp5348");
            register(ldc, "la-001", "windows-1252", "cp1252", "cp5348");
            register(ldc, "cs", "windows-1250", "cp1250", "cp5346");
            register(ldc, "ru-RU", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "es-AR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "cy", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-Hans-MO", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "az-Cyrl-AZ", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "fr-HT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "da", "windows-1252", "cp1252", "cp5348");
            register(ldc, "de", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nl-BQ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "yo-BJ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-MP", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-MO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-MT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-MS", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-MU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nl-BE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-MW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-MY", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-PS", "windows-1256", "cp1256");
            register(ldc, "pt-ST", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-NA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nl-AW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-NG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-NF", "windows-1252", "cp1252", "cp5348");
            register(ldc, "se-FI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "de-CH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "pap", "windows-1252", "cp1252", "cp5348");
            register(ldc, "de-BE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-QA", "windows-1256", "cp1256");
            register(ldc, "el", "cp5349", "windows-1253", "cp1253");
            register(ldc, "en", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-NR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-NU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-GA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es", "windows-1252", "cp1252", "cp5348");
            register(ldc, "et", "cp1257", "cp5353", "windows-1257");
            register(ldc, "eu", "windows-1252", "cp1252", "cp5348");
            register(ldc, "rw-RW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-NZ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tr-CY", "cp5350", "windows-1254", "cp1254");
            register(ldc, "hsb", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fa", "windows-1256", "cp1256");
            register(ldc, "fr-FR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "pt-TL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ff", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-KI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fi", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-KN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-SG_stroke", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "fo", "windows-1252", "cp1252", "cp5348");
            register(ldc, "moh-CA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sl-SI", "windows-1250", "cp1250", "cp5346");
            register(ldc, "ku-Arab", "windows-1256", "cp1256");
            register(ldc, "sah-RU", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "en-KY", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-CD", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fy", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-CG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-CF", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-CI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-CH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-LC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-CM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ga", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nl-CW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "gd", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-SA", "windows-1256", "cp1256");
            register(ldc, "de-DE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "gl", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-SD", "windows-1256", "cp1256");
            register(ldc, "fr-DZ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "gn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-HK_radstr", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "en-LS", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-LR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "uz-Cyrl-UZ", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "it-IT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ur-IN", "windows-1256", "cp1256");
            register(ldc, "uz-Latn", "cp5350", "windows-1254", "cp1254");
            register(ldc, "ar-SO", "windows-1256", "cp1256");
            register(ldc, "fr-DJ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-SS", "windows-1256", "cp1256");
            register(ldc, "ha", "windows-1252", "cp1252", "cp5348");
            register(ldc, "he", "windows-1255", "cp1255");
            register(ldc, "en-MH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-MG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-IE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-KW", "windows-1256", "cp1256");
            register(ldc, "ca-FR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "bin", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ha-Latn-GH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ms-MY", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-IN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-IM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "bin-NG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "gsw", "windows-1252", "cp1252", "cp5348");
            register(ldc, "hr", "windows-1250", "cp1250", "cp5346");
            register(ldc, "en-IO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "yo-NG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "hu", "windows-1250", "cp1250", "cp5346");
            register(ldc, "zh-Hans-HK", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "ar-LB", "windows-1256", "cp1256");
            register(ldc, "az-Latn-AZ", "cp5350", "windows-1254", "cp1254");
            register(ldc, "id", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-JE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ig", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-LY", "windows-1256", "cp1256");
            register(ldc, "en-JM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "is", "windows-1252", "cp1252", "cp5348");
            register(ldc, "it", "windows-1252", "cp1252", "cp5348");
            register(ldc, "iu", "windows-1252", "cp1252", "cp5348");
            register(ldc, "oc-FR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fy-NL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-CA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "se-SE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-029", "windows-1252", "cp1252", "cp5348");
            register(ldc, "smj-NO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sr-Cyrl-RS", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "pl-PL", "windows-1250", "cp1250", "cp5346");
            register(ldc, "fr-BF", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-BE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sd-Arab-PK", "windows-1256", "cp1256");
            register(ldc, "ja", "MS932", "Shift_JIS", "windows-932", "csWindows31J", "windows-31j");
            register(ldc, "fr-BJ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-BI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-BL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-MA", "windows-1256", "cp1256");
            register(ldc, "en-KE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ig-NG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-GD", "windows-1252", "cp1252", "cp5348");
            register(ldc, "hr-HR", "windows-1250", "cp1250", "cp5346");
            register(ldc, "en-GH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-GG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sma", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-GI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-MR", "windows-1256", "cp1256");
            register(ldc, "en-GM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "jv", "windows-1252", "cp1252", "cp5348");
            register(ldc, "de-AT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "is-IS", "windows-1252", "cp1252", "cp5348");
            register(ldc, "smj", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-GU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "smn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-GY", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sms", "windows-1252", "cp1252", "cp5348");
            register(ldc, "prs-AF", "windows-1256", "cp1256");
            register(ldc, "ro-MD", "windows-1250", "cp1250", "cp5346");
            register(ldc, "kl", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ko", "x-windows-949", "ms_949", "windows949", "EUC-KR", "ms949", "windows-949");
            register(ldc, "ca-IT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-HK", "windows-1252", "cp1252", "cp5348");
            register(ldc, "kr", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sr-Latn-ME", "windows-1250", "cp1250", "cp5346");
            register(ldc, "ku", "windows-1256", "cp1256");
            register(ldc, "sv-AX", "windows-1252", "cp1252", "cp5348");
            register(ldc, "jv-Java", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ky", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "ar-OM", "windows-1256", "cp1256");
            register(ldc, "en-001", "windows-1252", "cp1252", "cp5348");
            register(ldc, "gn-PY", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-001", "windows-1256", "cp1256");
            register(ldc, "la", "windows-1252", "cp1252", "cp5348");
            register(ldc, "lb", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-ID", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-419", "windows-1252", "cp1252", "cp5348");
            register(ldc, "smj-SE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "lt", "cp1257", "cp5353", "windows-1257");
            register(ldc, "quz-EC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "lv", "cp1257", "cp5353", "windows-1257");
            register(ldc, "zh-CN", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "ur-PK", "windows-1256", "cp1256");
            register(ldc, "en-ER", "windows-1252", "cp1252", "cp5348");
            register(ldc, "da-GL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ff-Latn-SN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fo-FO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "mk", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "fr-029", "windows-1252", "cp1252", "cp5348");
            register(ldc, "se-NO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "mn", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "jv-Latn-ID", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ms", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-FK", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-FJ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-FM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "mn-Cyrl", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "smn-FI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "dsb", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-IL", "windows-1256", "cp1256");
            register(ldc, "nb", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ro-RO", "windows-1250", "cp1250", "cp5346");
            register(ldc, "it-CH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-SG_phoneb", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "en-GB", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nl", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "no", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-CA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-CC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "lt-LT", "cp1257", "cp5353", "windows-1257");
            register(ldc, "en-CK", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-IQ", "windows-1256", "cp1256");
            register(ldc, "en-CM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tn-ZA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "oc", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-CX", "windows-1252", "cp1252", "cp5348");
            register(ldc, "quz-BO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-US", "windows-1252", "cp1252", "cp5348");
            register(ldc, "af-NA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tzm-Latn-DZ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ca-ES", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sq-MK", "windows-1250", "cp1250", "cp5346");
            register(ldc, "es-UY", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-CN_stroke", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "ar-JO", "windows-1256", "cp1256");
            register(ldc, "az-Latn", "cp5350", "windows-1254", "cp1254");
            register(ldc, "es-VE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-DM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sr-Latn", "windows-1250", "cp1250", "cp5346");
            register(ldc, "xh-ZA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ms-SG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-ES_tradnl", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-KM", "windows-1256", "cp1256");
            register(ldc, "pl", "windows-1250", "cp1250", "cp5346");
            register(ldc, "jv-Latn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "pt", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-AG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "hr-BA", "windows-1250", "cp1250", "cp5346");
            register(ldc, "uz-Latn-UZ", "cp5350", "windows-1254", "cp1254");
            register(ldc, "en-AI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "el-GR", "cp5349", "windows-1253", "cp1253");
            register(ldc, "sr-Latn-RS", "windows-1250", "cp1250", "cp5346");
            register(ldc, "zh-Hans", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "gd-GB", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nl-NL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-Hant", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "kr-Latn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-AS", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-AU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "he-IL", "windows-1255", "cp1255");
            register(ldc, "mk-MK", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "tn-BW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tg-Cyrl-TJ", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "en-BB", "windows-1252", "cp1252", "cp5348");
            register(ldc, "arn-CL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-BE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-SV", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tt-RU", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "ar-DJ", "windows-1256", "cp1256");
            register(ldc, "gsw-CH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-BM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-BS", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-EG", "windows-1256", "cp1256");
            register(ldc, "quc-Latn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-BW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-DZ", "windows-1256", "cp1256");
            register(ldc, "hu-HU_technl", "windows-1250", "cp1250", "cp5346");
            register(ldc, "fil-PH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "rm", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-BZ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ro", "windows-1250", "cp1250", "cp5346");
            register(ldc, "nb-NO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-ER", "windows-1256", "cp1256");
            register(ldc, "jv-Java-ID", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ru", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "es-PR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "rw", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-MO", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "sq-XK", "windows-1250", "cp1250", "cp5346");
            register(ldc, "es-PY", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sd", "windows-1256", "cp1256");
            register(ldc, "zh-MO_radstr", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "se", "windows-1252", "cp1252", "cp5348");
            register(ldc, "pa-Arab", "windows-1256", "cp1256");
            register(ldc, "sk", "windows-1250", "cp1250", "cp5346");
            register(ldc, "sl", "windows-1250", "cp1250", "cp5346");
            register(ldc, "arn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ru-BY", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "sq", "windows-1250", "cp1250", "cp5346");
            register(ldc, "sr", "windows-1250", "cp1250", "cp5346");
            register(ldc, "iu-Latn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sv", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sw", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ibb", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tzm-Latn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ha-Latn-NE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ca-AD", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-YT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sma-NO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ha-Latn-NG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tg", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "th", "windows-874", "ms-874", "x-windows-874", "ms874", "ISO-8859-11");
            register(ldc, "tk", "windows-1250", "cp1250", "cp5346");
            register(ldc, "da-DK", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tr", "cp5350", "windows-1254", "cp1254");
            register(ldc, "gsw-FR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tt", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "pt-CV", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sd-Arab", "windows-1256", "cp1256");
            register(ldc, "lv-LV", "cp1257", "cp5353", "windows-1257");
            register(ldc, "sr-Cyrl-ME", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "ha-Latn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ug", "windows-1256", "cp1256");
            register(ldc, "uk", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "el-CY", "cp5349", "windows-1253", "cp1253");
            register(ldc, "ur", "windows-1256", "cp1256");
            register(ldc, "haw", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-WF", "windows-1252", "cp1252", "cp5348");
            register(ldc, "uz", "cp5350", "windows-1254", "cp1254");
            register(ldc, "tzm", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-TW_radstr", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "hsb-DE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nn-NO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-PE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "prs", "windows-1256", "cp1256");
            register(ldc, "fa-IR", "windows-1256", "cp1256");
            register(ldc, "fr-VU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-PH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "vi", "cp1258", "windows-1258");
            register(ldc, "es-PA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nso-ZA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tr-TR", "cp5350", "windows-1254", "cp1254");
            register(ldc, "pa-Arab-PK", "windows-1256", "cp1256");
            register(ldc, "vi-VN", "cp1258", "windows-1258");
            register(ldc, "quc", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-TD", "windows-1252", "cp1252", "cp5348");
            register(ldc, "bs-Cyrl-BA", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "kl-GL", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-TG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ff-Latn-MR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "pt-AO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-TN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-AE", "windows-1256", "cp1256");
            register(ldc, "fr-SY", "windows-1252", "cp1252", "cp5348");
            register(ldc, "quz", "windows-1252", "cp1252", "cp5348");
            register(ldc, "wo", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tg-Cyrl", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "sw-CD", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zh-HK", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "pt-BR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sv-FI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-BH", "windows-1256", "cp1256");
            register(ldc, "nl-SX", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-MX", "windows-1252", "cp1252", "cp5348");
            register(ldc, "br-FR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "xh", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nl-SR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sma-SE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "id-ID", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ff-Latn-NG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-NI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "th-TH", "windows-874", "ms-874", "x-windows-874", "ms874", "ISO-8859-11");
            register(ldc, "ba-RU", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "de-DE_phoneb", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-RE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "pt-GW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "moh", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ff-Latn-GN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-ZA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-ZM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "yo", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tk-TM", "windows-1250", "cp1250", "cp5346");
            register(ldc, "ru-MD", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "en-ZW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-SC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ibb-NG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "kr-Latn-NG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ug-CN", "windows-1256", "cp1256");
            register(ldc, "haw-US", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nb-SJ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-SN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "eu-ES", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sr-Cyrl-BA", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "zh", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "quz-PE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "x-IV_mathan", "windows-1252", "cp1252", "cp5348");
            register(ldc, "gsw-LI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "tzm-Arab-MA", "windows-1256", "cp1256");
            register(ldc, "fr-RW", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ga-IE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "zu", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ja-JP_radstr", "MS932", "Shift_JIS", "windows-932", "csWindows31J", "windows-31j");
            register(ldc, "en-WS", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-HN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ru-KZ", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "fr-PF", "windows-1252", "cp1252", "cp5348");
            register(ldc, "gl-ES", "windows-1252", "cp1252", "cp5348");
            register(ldc, "nso", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-PM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "rm-CH", "windows-1252", "cp1252", "cp5348");
            register(ldc, "bg-BG", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "ru-KG", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "cs-CZ", "windows-1250", "cp1250", "cp5346");
            register(ldc, "lb-LU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "bs-Latn-BA", "windows-1250", "cp1250", "cp5346");
            register(ldc, "zh-TW", "windows-950", "x-windows-950", "Big5", "ms950");
            register(ldc, "sk-SK", "windows-1250", "cp1250", "cp5346");
            register(ldc, "sr-Cyrl", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "sq-AL", "windows-1250", "cp1250", "cp5346");
            register(ldc, "ca-ES-valencia", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sv-SE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "uz-Cyrl", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "mn-MN", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "uk-UA", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "en-US", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-NC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-NE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ff-Latn-CM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-VC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-ML", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-VG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-MQ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-VI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-MR", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-MU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "quc-Latn-GT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ja-JP", "MS932", "Shift_JIS", "windows-932", "csWindows31J", "windows-31j");
            register(ldc, "en-VU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ar-YE", "windows-1256", "cp1256");
            register(ldc, "es-GT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-GQ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "wo-SN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "be-BY", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "zh-SG", "GBK", "CP936", "windows-936", "GB18030");
            register(ldc, "bs-Cyrl", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "et-EE", "cp1257", "cp5353", "windows-1257");
            register(ldc, "fi-FI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "it-SM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sr-Latn-XK", "windows-1250", "cp1250", "cp5346");
            register(ldc, "tzm-Arab", "windows-1256", "cp1256");
            register(ldc, "en-SS", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sms-FI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-SX", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-SZ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-DO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-TC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-KM", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ff-Latn", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-EC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-TK", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sah", "windows-1251", "cp1251", "cp5347", "ansi-1251");
            register(ldc, "en-TO", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-MA", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-TT", "windows-1252", "cp1252", "cp5348");
            register(ldc, "de-LI", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-MC", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-TV", "windows-1252", "cp1252", "cp5348");
            register(ldc, "ms-BN", "windows-1252", "cp1252", "cp5348");
            register(ldc, "sw-KE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "es-ES", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-MG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-TZ", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-MF", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-UG", "windows-1252", "cp1252", "cp5348");
            register(ldc, "de-LU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "fr-LU", "windows-1252", "cp1252", "cp5348");
            register(ldc, "dsb-DE", "windows-1252", "cp1252", "cp5348");
            register(ldc, "en-UM", "windows-1252", "cp1252", "cp5348");

            localeDefaultCharsets = Collections.unmodifiableMap(ldc);
        }

        private static void register(Map<String, Set<String>> dst, String locale, String... charsetNames) {
            assert dst != null;

            for (String charsetName : charsetNames) {
                dst.computeIfAbsent(locale, k -> new HashSet<>()).add(charsetName);
            }
        }

        @NotNull
        static Set<String> getByLocale(Locale locale) {
            Set<String> codePages;

            if (locale == null) {
                codePages = localeDefaultCharsets.get("");
            } else {
                final String languageTag = locale.toLanguageTag();
                if (localeDefaultCharsets.containsKey(languageTag)) {
                    codePages = localeDefaultCharsets.get(languageTag);
                } else {
                    codePages = localeDefaultCharsets.get("");
                }
            }

            return codePages;
        }

    }

    private static class CharsetCandidate {

        private String charsetName;

        private int score;

        private List<String> sources;

    }

}
