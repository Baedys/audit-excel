using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    internal static class ArixcelMapProfile
    {
        const int MaxBytes = 65536;
        static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);
        static readonly FormulaMapKind[] Kinds = Enum.GetValues(typeof(FormulaMapKind)).Cast<FormulaMapKind>()
            .Where(kind => kind != FormulaMapKind.Unknown).ToArray();
        static readonly Dictionary<string, FormulaMapKind> KindNames =
            Kinds.ToDictionary(kind => kind.ToString(), StringComparer.Ordinal);

        internal static HighlightPalette Load(string path, HighlightPalette baseline)
        {
            if (String.IsNullOrWhiteSpace(path)) throw new ArgumentException("An explicit profile file path is required.", "path");
            if (baseline == null) throw new ArgumentNullException("baseline");
            var root = RequireObject(new ProfileJson(ReadText(path)).Read(), "profile");
            HighlightPalette palette = baseline.Clone();
            object legacyMarker;
            if (root.TryGetValue("UniqueFormula", out legacyMarker) && legacyMarker != null)
            {
                ApplyColors(palette, root);
                return palette;
            }

            object profile;
            root.TryGetValue("FormulaMaps", out profile);
            var maps = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, object> item in RequireObject(profile, "FormulaMaps"))
            {
                if (!String.Equals(item.Key, "AnalyseUFs", StringComparison.OrdinalIgnoreCase) &&
                    !String.Equals(item.Key, "Colors", StringComparison.OrdinalIgnoreCase) &&
                    !String.Equals(item.Key, "CustomColors", StringComparison.OrdinalIgnoreCase)) continue;
                if (maps.ContainsKey(item.Key)) throw new FormatException("Duplicate FormulaMaps property: " + item.Key + ".");
                maps.Add(item.Key, item.Value);
            }
            object flag, colors, swatches;
            if (maps.TryGetValue("AnalyseUFs", out flag) && !(flag is bool))
                throw new FormatException("AnalyseUFs must be true or false.");
            palette.AnalyzeUniqueFormulas = flag == null || (bool)flag;
            maps.TryGetValue("Colors", out colors);
            ApplyColors(palette, RequireObject(colors, "Colors"));
            maps.TryGetValue("CustomColors", out swatches);
            var list = swatches as List<object>;
            if (list == null) throw new FormatException("CustomColors must be a JSON array.");
            if (list.Count > PaletteStorage.MaxCustomColors) throw new FormatException("Too many custom colors in the palette.");
            palette.CustomColors.Clear();
            foreach (object color in list) palette.CustomColors.Add(ReadColor(color));
            return palette;
        }

        internal static void Save(string path, HighlightPalette palette)
        {
            if (palette == null) throw new ArgumentNullException("palette");
            HighlightPalette snapshot = palette.Clone();
            int?[] customColors = PaletteStorage.SnapshotCustomColors(snapshot.CustomColors);
            // All JSON strings here are fixed identifiers or formatted RGB values.
            PaletteStorage.WriteAtomically(path, stream =>
            {
                using (var writer = new StreamWriter(stream, Utf8, 4096, true))
                {
                    writer.NewLine = "\n";
                    writer.WriteLine("{");
                    writer.WriteLine("  \"File\": \"Arixcel Explorer Settings\",");
                    writer.WriteLine("  \"Version\": \"9.0.9623\",");
                    writer.WriteLine("  \"FormulaMaps\": {");
                    writer.WriteLine("    \"AnalyseUFs\": " + (snapshot.AnalyzeUniqueFormulas ? "true" : "false") + ",");
                    writer.WriteLine("    \"Colors\": {");
                    writer.WriteLine(String.Join(",\n", Kinds.Select(kind =>
                    {
                        HighlightStyle style = snapshot.Entries[kind];
                        return "      \"" + kind + "\": " + WriteColor(style.Enabled ? style.FillRgb : null);
                    })));
                    writer.WriteLine("    },");
                    writer.WriteLine("    \"CustomColors\": [" + String.Join(", ", customColors.Select(WriteColor)) + "]");
                    writer.WriteLine("  }");
                    writer.WriteLine("}");
                }
            });
        }

        static string ReadText(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (stream.Length > MaxBytes) throw new FormatException("The profile exceeds 64 KiB.");
                byte[] bytes = new byte[MaxBytes + 1];
                int size = 0;
                while (size < bytes.Length)
                {
                    int count = stream.Read(bytes, size, bytes.Length - size);
                    if (count == 0) break;
                    size += count;
                }
                if (size > MaxBytes) throw new FormatException("The profile exceeds 64 KiB.");
                int start = size >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
                try { return Utf8.GetString(bytes, start, size - start); }
                catch (DecoderFallbackException ex) { throw new FormatException("The profile is not valid UTF-8.", ex); }
            }
        }

        static Dictionary<string, object> RequireObject(object value, string name)
        {
            var result = value as Dictionary<string, object>;
            if (result == null) throw new FormatException(name + " must be a JSON object.");
            return result;
        }

        static void ApplyColors(HighlightPalette palette, Dictionary<string, object> colors)
        {
            foreach (KeyValuePair<string, object> item in colors)
            {
                FormulaMapKind kind;
                if (!KindNames.TryGetValue(item.Key, out kind)) throw new FormatException("Unknown profile color category: " + item.Key + ".");
                HighlightStyle style = palette.Entries[kind];
                style.FillRgb = ReadColor(item.Value);
                if (style.FillRgb.HasValue) style.Enabled = true;
            }
        }

        static int? ReadColor(object value)
        {
            var text = value as string;
            if (text == null) throw new FormatException("Profile colors must be JSON strings.");
            if (text.Length == 0) return null;
            Color color;
            try { color = ColorTranslator.FromHtml(text); }
            catch (Exception)
            {
                try { color = (Color)TypeDescriptor.GetConverter(typeof(Color)).ConvertFromInvariantString(text); }
                catch (Exception) { return null; } // Framework color converters can throw a plain Exception for bad text.
            }
            return color.IsEmpty || color.A != 255 ? (int?)null : color.ToArgb() & 0xFFFFFF;
        }

        static string WriteColor(int? value)
        {
            return value.HasValue ? "\"#" + value.Value.ToString("X6", CultureInfo.InvariantCulture) + "\"" : "\"\"";
        }

        // Framework 4's JSON reader accepts trailing commas and ignores trailing documents.
        // Keep strict JSON tokens and bounded nesting without adding an add-in dependency.
        sealed class ProfileJson
        {
            static readonly Regex Token = new Regex(
                @"\G[ \t\r\n]*(?<token>[{}\[\]:,]|true|false|null|-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?|""(?>(?:[^""\\\x00-\x1F]+|\\(?:[""\\/bfnrt]|u[0-9a-fA-F]{4}))*)""|\z)");
            static readonly Regex Escape = new Regex(@"\\(u[0-9a-fA-F]{4}|.)");
            static readonly object Number = new object();
            readonly string source;
            int position;
            string token;

            internal ProfileJson(string source) { this.source = source; }

            internal object Read()
            {
                Advance();
                object value = ReadValue(0);
                if (token.Length != 0) throw InvalidJson();
                return value;
            }

            object ReadValue(int depth)
            {
                if (depth > 12) throw new FormatException("The profile is nested too deeply.");
                if (Take("{"))
                {
                    var result = new Dictionary<string, object>(StringComparer.Ordinal);
                    if (Take("}")) return result;
                    do
                    {
                        string key = ReadString();
                        Require(":");
                        if (result.ContainsKey(key)) throw new FormatException("Duplicate JSON property: " + key + ".");
                        result.Add(key, ReadValue(depth + 1));
                    } while (Take(","));
                    Require("}");
                    return result;
                }
                if (Take("["))
                {
                    var result = new List<object>();
                    if (Take("]")) return result;
                    do { result.Add(ReadValue(depth + 1)); } while (Take(","));
                    Require("]");
                    return result;
                }
                if (Take("null")) return null;
                if (Take("true")) return true;
                if (Take("false")) return false;
                if (token.Length == 0) throw InvalidJson();
                if (token[0] == '"') return ReadString();
                if (token[0] != '-' && (token[0] < '0' || token[0] > '9')) throw InvalidJson();
                Advance();
                return Number;
            }

            string ReadString()
            {
                if (token.Length == 0 || token[0] != '"') throw InvalidJson();
                string value = Escape.Replace(token.Substring(1, token.Length - 2), match =>
                {
                    char c = match.Value[1];
                    switch (c)
                    {
                        case 'b': return "\b";
                        case 'f': return "\f";
                        case 'n': return "\n";
                        case 'r': return "\r";
                        case 't': return "\t";
                        case 'u':
                            return ((char)Int32.Parse(match.Value.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString();
                        default: return c.ToString();
                    }
                });
                Advance();
                return value;
            }

            bool Take(string expected)
            {
                if (token != expected) return false;
                Advance();
                return true;
            }
            void Require(string expected)
            {
                if (!Take(expected)) throw InvalidJson();
            }
            void Advance()
            {
                Match match = Token.Match(source, position);
                if (!match.Success) throw InvalidJson();
                token = match.Groups["token"].Value;
                position += match.Length;
            }
            static FormatException InvalidJson() { return new FormatException("Invalid map profile JSON."); }
        }
    }
}
