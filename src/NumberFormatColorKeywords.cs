using System;
using System.Collections.Generic;
using System.Globalization;

namespace FormulaExplorer
{
    internal static class NumberFormatColorKeywords
    {
        // Reserved Excel number-format words, ordered Black/Blue/Cyan/Green/Magenta/Red/White/Yellow.
        static readonly string[] English = { "black", "blue", "cyan", "green", "magenta", "red", "white", "yellow" };
        static readonly Dictionary<string, string[]> Words = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase) {
            { "bg-BG", new[] { "\u0447\u0435\u0440\u043d\u043e", "\u0441\u0438\u043d\u044c\u043e", "\u0446\u0438\u0430\u043d", "\u0437\u0435\u043b\u0435\u043d\u043e", "\u043c\u0430\u0433\u0435\u043d\u0442\u0430", "\u0447\u0435\u0440\u0432\u0435\u043d\u043e", "\u0431\u044f\u043b\u043e", "\u0436\u044a\u043b\u0442\u043e" } },
            { "cs-CZ", new[] { "\u010dern\u00e1", "modr\u00e1", "azurov\u00e1", "zelen\u00e1", "purpurov\u00e1", "\u010derven\u00e1", "b\u00edl\u00e1", "\u017elut\u00e1" } },
            { "da-DK", new[] { "sort", "bl\u00e5", "akvamarin", "lysegr\u00f8n", "lyslilla", "r\u00f8d", "hvid", "gul" } },
            { "de-DE", new[] { "schwarz", "blau", "zyan", "gr\u00fcn", "magenta", "rot", "wei\u00df", "gelb" } },
            { "el-GR", new[] { "\u03bc\u03b1\u03cd\u03c1\u03bf", "\u03bc\u03c0\u03bb\u03b5", "\u03b3\u03b1\u03bb\u03ac\u03b6\u03b9\u03bf", "\u03c0\u03c1\u03ac\u03c3\u03b9\u03bd\u03bf", "\u03bc\u03b1\u03c4\u03b6\u03ad\u03bd\u03c4\u03b1", "\u03ba\u03cc\u03ba\u03ba\u03b9\u03bd\u03bf", "\u03bb\u03b5\u03c5\u03ba\u03cc", "\u03ba\u03af\u03c4\u03c1\u03b9\u03bd\u03bf" } },
            { "en-US", new[] { "black", "blue", "cyan", "green", "magenta", "red", "white", "yellow" } },
            { "es-ES", new[] { "negro", "azul", "cian", "verde", "magenta", "rojo", "blanco", "amarillo" } },
            { "et-EE", new[] { "must", "sinine", "ts\u00fcaan", "roheline", "magenta", "punane", "valge", "kollane" } },
            { "fi-FI", new[] { "musta", "sininen", "syaani", "vihre\u00e4", "magenta", "punainen", "valkoinen", "keltainen" } },
            { "fr-FR", new[] { "noir", "bleu", "cyan", "vert", "magenta", "rouge", "blanc", "jaune" } },
            { "hr-HR", new[] { "crna", "plava", "cijan", "zelena", "magenta", "crvena", "bijela", "\u017euta" } },
            { "hu-HU", new[] { "fekete", "k\u00e9k", "ci\u00e1nk\u00e9k", "z\u00f6ld", "b\u00edbor", "piros", "feh\u00e9r", "s\u00e1rga" } },
            { "id-ID", new[] { "hitam", "biru", "sian", "hijau", "magenta", "merah", "putih", "kuning" } },
            { "it-IT", new[] { "nero", "blu", "ciano", "verde", "magenta", "rosso", "bianco", "giallo" } },
            { "lt-LT", new[] { "juoda", "m\u0117lyna", "\u017ealsvai m\u0117lyna", "\u017ealia", "rausvai raudona", "raudona", "balta", "geltona" } },
            { "lv-LV", new[] { "melna", "zila", "ci\u0101na", "za\u013ca", "mad\u017eentas", "sarkana", "balta", "dzeltena" } },
            { "nb-NO", new[] { "svart", "bl\u00e5", "cyan", "gr\u00f8nn", "magenta", "r\u00f8d", "hvitt", "gule" } },
            { "nl-NL", new[] { "zwart", "blauw", "cyaan", "groen", "magenta", "rood", "wit", "geel" } },
            { "pl-PL", new[] { "czarny", "niebieski", "b\u0142\u0119kitny", "zielony", "amarantowy", "czerwony", "bia\u0142y", "\u017c\u00f3\u0142ty" } },
            { "pt-BR", new[] { "preto", "azul", "ciano", "verde", "magenta", "vermelho", "branco", "amarelo" } },
            { "pt-PT", new[] { "preto", "azul", "turquesa", "verde", "magenta", "vermelho", "branco", "amarelo" } },
            { "ro-RO", new[] { "negru", "albastru", "bleu", "verde", "fucsia", "ro\u0219u", "alb", "galben" } },
            { "ru-RU", new[] { "\u0447\u0435\u0440\u043d\u044b\u0439", "\u0441\u0438\u043d\u0438\u0439", "\u0433\u043e\u043b\u0443\u0431\u043e\u0439", "\u0437\u0435\u043b\u0435\u043d\u044b\u0439", "\u0444\u0438\u043e\u043b\u0435\u0442\u043e\u0432\u044b\u0439", "\u043a\u0440\u0430\u0441\u043d\u044b\u0439", "\u0431\u0435\u043b\u044b\u0439", "\u0436\u0435\u043b\u0442\u044b\u0439" } },
            { "sk-SK", new[] { "\u010dierna", "modr\u00e1", "az\u00farov\u00e1", "zelen\u00e1", "purpurov\u00e1", "\u010derven\u00e1", "biela", "\u017elt\u00e1" } },
            { "sl-SI", new[] { "\u010drna", "modra", "cijan", "zelena", "magenta", "rde\u010da", "bela", "rumena" } },
            { "sr-Latn-RS", new[] { "crna", "plava", "cijan", "zelena", "magenta", "crvena", "bela", "\u017euta" } },
            { "sv-SE", new[] { "svart", "bl\u00e5", "cyanbl\u00e5", "gr\u00f6n", "magenta", "r\u00f6d", "vit", "gul" } },
            { "tr-TR", new[] { "siyah", "mavi", "deniz mavisi", "ye\u015fil", "eflatun", "k\u0131rm\u0131z\u0131", "beyaz", "sar\u0131" } },
            { "uk-UA", new[] { "\u0447\u043e\u0440\u043d\u0438\u0439", "\u0441\u0438\u043d\u0456\u0439", "\u0431\u043b\u0430\u043a\u0438\u0442\u043d\u0438\u0439", "\u0437\u0435\u043b\u0435\u043d\u0438\u0439", "\u043c\u0430\u043b\u0438\u043d\u043e\u0432\u0438\u0439", "\u0447\u0435\u0440\u0432\u043e\u043d\u0438\u0439", "\u0431\u0456\u043b\u0438\u0439", "\u0436\u043e\u0432\u0442\u0438\u0439" } },
        };

        internal static string Translate(string value, string language, CultureInfo culture)
        {
            string[] words;
            if (String.IsNullOrEmpty(value) || !Words.TryGetValue(language, out words)) return value;
            string prefix = "", suffix = "", key = value;
            if (key.Length > 2 && ((key[0] == '<' && key[key.Length - 1] == '>') || (key[0] == '[' && key[key.Length - 1] == ']')))
            { prefix = key.Substring(0, 1); suffix = key.Substring(key.Length - 1); key = key.Substring(1, key.Length - 2); }
            bool period = key.EndsWith(".", StringComparison.Ordinal);
            if (period) key = key.Substring(0, key.Length - 1);
            int index = Array.FindIndex(English, word => String.Equals(word, key, StringComparison.OrdinalIgnoreCase));
            if (index < 0) return value;
            string result = words[index];
            if (key == key.ToUpperInvariant()) result = result.ToUpper(culture);
            else if (Char.IsUpper(key[0])) result = Char.ToUpper(result[0], culture) + result.Substring(1);
            return prefix + result + (period ? "." : "") + suffix;
        }
    }
}
