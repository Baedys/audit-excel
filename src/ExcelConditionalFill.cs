using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace FormulaExplorer
{
    // Owns its rule RCW and an IUnknown ticket. The session proves worksheet
    // and rule membership before every mutation; names are never live identities.
    public sealed class ExcelConditionalFill : IDisposable
    {
        object rule;
        IntPtr identity;
        string fullWitness, relativeWitness;
        internal object Rule { get { Alive(); return rule; } }
        internal long Identity { get { Alive(); return identity.ToInt64(); } }
        public ExcelConditionalFill(object ownedRule)
        {
            if (ownedRule == null || !Marshal.IsComObject(ownedRule)) throw new ArgumentException("Règle conditionnelle COM attendue.", "ownedRule");
            identity = Marshal.GetIUnknownForObject(ownedRule); rule = ownedRule;
        }
        ~ExcelConditionalFill() { ReleaseIdentity(); }
        void Alive() { if (rule == null) throw new ObjectDisposedException("ExcelConditionalFill"); }
        static Exception Unwrap(Exception error)
        { return error is TargetInvocationException && error.InnerException != null ? error.InnerException : error; }
        internal static object Get(object target, string property)
        {
            try { return target.GetType().InvokeMember(property, BindingFlags.GetProperty, null, target, null); }
            catch (TargetInvocationException error) { throw Unwrap(error); }
        }
        static void Set(object target, string property, object value)
        {
            try { target.GetType().InvokeMember(property, BindingFlags.SetProperty, null, target, new[] { value }); }
            catch (TargetInvocationException error) { throw Unwrap(error); }
        }
        internal static bool MissingProperty(Exception error)
        {
            uint code = (uint)Unwrap(error).HResult;
            return code == 0x80020003 || code == 0x80020006 || code == 0x800A03EC || code == 0x80070057 || code == 0x8002000B;
        }
        public ConditionalFillState Read()
        {
            Alive(); object fill = null;
            try
            {
                try { fill = Get(rule, "Interior"); }
                catch (Exception error) { if (MissingProperty(error)) throw new NotSupportedException("Cette famille de règle n'expose pas de fond conditionnel.", error); throw; }
                if (fill == null || !Marshal.IsComObject(fill)) throw new NotSupportedException("Cette famille de règle n'expose pas de fond conditionnel.");
                var values = new List<KeyValuePair<string, ConditionalFillScalar>>();
                foreach (string field in ConditionalFillState.FieldNames)
                {
                    ConditionalFillScalar scalar;
                    try
                    {
                        object value = Get(fill, field);
                        if (value == null || value == DBNull.Value) scalar = ConditionalFillScalar.Null;
                        else if (value is byte || value is short || value is int || value is long || value is float || value is double || value is decimal)
                            scalar = ConditionalFillScalar.Number(Convert.ToDouble(value, CultureInfo.InvariantCulture));
                        else throw new NotSupportedException("Propriété de fond conditionnel non scalaire : " + field);
                    }
                    catch (Exception error)
                    {
                        if ((field == "ThemeColor" || field == "PatternThemeColor") && MissingProperty(error)) scalar = ConditionalFillScalar.Unavailable;
                        else throw;
                    }
                    values.Add(new KeyValuePair<string, ConditionalFillScalar>(field, scalar));
                }
                var pattern = values.Find(pair => pair.Key == "Pattern").Value;
                if (Numeric(pattern, 4000) || Numeric(pattern, 4001))
                    throw new NotSupportedException("Les dégradés de fond conditionnel exigent un journal de points de couleur.");
                return new ConditionalFillState(values);
            }
            finally { ExcelBridge.Release(fill); }
        }
        public void Mask()
        {
            Alive(); ExcelConditionalFillContext.Run(rule, delegate
            {
                object fill = null;
                try { fill = Get(rule, "Interior"); Set(fill, "ColorIndex", DBNull.Value); }
                finally { ExcelBridge.Release(fill); }
            });
        }
        static bool Numeric(ConditionalFillScalar value, double expected)
        { return value.Kind == ConditionalFillScalarKind.Number && value.NumberValue == expected; }
        static bool Theme(ConditionalFillScalar value)
        { return value.Kind == ConditionalFillScalarKind.Number && value.NumberValue >= 1 && value.NumberValue <= 12; }
        public void Restore(ConditionalFillState state)
        {
            if (state == null) throw new ArgumentNullException("state");
            Alive(); if (Read().Equals(state)) return;
            ExcelConditionalFillContext.Run(rule, delegate { RestoreCore(state); });
        }
        void RestoreCore(ConditionalFillState state)
        {
            object fill = null;
            try
            {
                fill = Get(rule, "Interior");
                ConditionalFillState actual = Read();
                ConditionalFillScalar index = state.Get("ColorIndex"), theme = state.Get("ThemeColor");
                // Setting diagnostics such as ThemeColor=0/-1 or a palette index
                // would destroy or quantize RGB. Only write the defining color.
                if (!actual.Get("ColorIndex").Equals(index) || !actual.Get("Color").Equals(state.Get("Color")) || !actual.Get("ThemeColor").Equals(theme))
                {
                    if (Numeric(index, -4142) || Numeric(index, -4105)) Set(fill, "ColorIndex", (int)index.NumberValue);
                    else if (Theme(theme)) Set(fill, "ThemeColor", (int)theme.NumberValue);
                    else if (state.Get("Color").Kind == ConditionalFillScalarKind.Number) Set(fill, "Color", state.Get("Color").NumberValue);
                    else Set(fill, "ColorIndex", DBNull.Value);
                }
                actual = Read();
                if (!actual.Get("TintAndShade").Equals(state.Get("TintAndShade")))
                    SetScalar(fill, "TintAndShade", state.Get("TintAndShade"));
                actual = Read();
                if (!actual.Get("PatternColorIndex").Equals(state.Get("PatternColorIndex")) ||
                    !actual.Get("PatternColor").Equals(state.Get("PatternColor")) ||
                    !actual.Get("PatternThemeColor").Equals(state.Get("PatternThemeColor")))
                {
                    ConditionalFillScalar patternIndex = state.Get("PatternColorIndex"), patternTheme = state.Get("PatternThemeColor");
                    if (Numeric(patternIndex, -4142) || Numeric(patternIndex, -4105)) Set(fill, "PatternColorIndex", (int)patternIndex.NumberValue);
                    else if (Theme(patternTheme)) Set(fill, "PatternThemeColor", (int)patternTheme.NumberValue);
                    else if (state.Get("PatternColor").Kind == ConditionalFillScalarKind.Number) Set(fill, "PatternColor", state.Get("PatternColor").NumberValue);
                    else Set(fill, "PatternColorIndex", DBNull.Value);
                }
                actual = Read();
                if (!actual.Get("PatternTintAndShade").Equals(state.Get("PatternTintAndShade")))
                    SetScalar(fill, "PatternTintAndShade", state.Get("PatternTintAndShade"));
                actual = Read();
                if (!actual.Get("Pattern").Equals(state.Get("Pattern"))) SetScalar(fill, "Pattern", state.Get("Pattern"));
                if (!Read().Equals(state)) throw new InvalidOperationException("Le fond conditionnel n'a pas retrouvé toutes ses propriétés. Son historique reste conservé.");
            }
            finally { ExcelBridge.Release(fill); }
        }
        static void SetScalar(object fill, string field, ConditionalFillScalar value)
        {
            if (value.Kind == ConditionalFillScalarKind.Unavailable) return;
            Set(fill, field, value.Kind == ConditionalFillScalarKind.Null ? (object)DBNull.Value : value.NumberValue);
        }
        static string Text(object value)
        {
            if (value == null || value == DBNull.Value) return "null";
            if (value is bool) return (bool)value ? "true" : "false";
            if (value is string) return "text:" + (string)value;
            return "number:" + Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        static void Witness(StringBuilder text, object value, string name)
        {
            string result;
            try { result = Text(Get(value, name)); }
            catch (Exception error) { if (!MissingProperty(error)) throw; result = "unavailable"; }
            text.Append(name.Length).Append(':').Append(name).Append('=').Append(result.Length).Append(':').Append(result).Append(';');
        }
        static void ReferenceFlags(ExpressionNode node, StringBuilder flags)
        {
            if (node.Kind == "Range" && node.Children.Count == 2)
            {
                string left = node.Children[0].Text.Substring(node.Children[0].Text.LastIndexOf('!') + 1).Trim();
                string right = node.Children[1].Text.Substring(node.Children[1].Text.LastIndexOf('!') + 1).Trim();
                bool columns = Regex.IsMatch(left, @"(?i)^\$?[A-Z]{1,3}$") && Regex.IsMatch(right, @"(?i)^\$?[A-Z]{1,3}$");
                bool rows = Regex.IsMatch(left, @"^\$?[1-9][0-9]*$") && Regex.IsMatch(right, @"^\$?[1-9][0-9]*$");
                if (columns || rows)
                    flags.Append(columns ? 'C' : 'R').Append(left.StartsWith("$", StringComparison.Ordinal) ? 'a' : 'r').Append(right.StartsWith("$", StringComparison.Ordinal) ? 'a' : 'r');
            }
            if (node.Kind == "Reference")
            {
                string local = node.Text.Substring(node.Text.LastIndexOf('!') + 1);
                foreach (Match match in Regex.Matches(local, @"(?i)(\$?)[A-Z]{1,3}(\$?)[0-9]+"))
                    flags.Append(match.Groups[1].Length == 0 ? 'r' : 'a').Append(match.Groups[2].Length == 0 ? 'r' : 'a');
            }
            foreach (ExpressionNode child in node.Children) ReferenceFlags(child, flags);
        }
        internal static string FormulaReferenceModes(string formula)
        { var flags = new StringBuilder(); ReferenceFlags(FormulaSyntax.Parse(formula), flags); return flags.ToString(); }
        string FormulaWitness(string name, object applies, bool relative)
        {
            object value;
            try { value = Get(rule, name); }
            catch (Exception error) { if (!MissingProperty(error)) throw; return "unavailable"; }
            string formula = value as string;
            if (!relative || formula == null || !formula.StartsWith("=", StringComparison.Ordinal)) return Text(value);
            object app = null, cells = null, anchor = null;
            try
            {
                app = Get(rule, "Application"); cells = Get(applies, "Cells"); anchor = ((dynamic)cells).Item[1, 1];
                // Relative conversion follows a retained AppliesTo anchor when
                // rows move. Preserve reference dollar modes as extra witnesses.
                string normalized = (string)((dynamic)app).ConvertFormula(formula, 1, -4150, 4, anchor);
                return "relative:" + FormulaReferenceModes(formula) + ":" + normalized;
            }
            catch (FormatException) { return Text(value); }
            catch (Exception error) { if (!MissingProperty(error)) throw; return Text(value); }
            finally { ExcelBridge.Release(anchor); ExcelBridge.Release(cells); ExcelBridge.Release(app); }
        }
        // Signature excludes fill and priority. Unique rules can be reordered;
        // ambiguous identical cohorts require an explicit conservative policy.
        internal string Signature(bool relative = false)
        {
            Alive();
            if (fullWitness != null) return relative ? relativeWitness : fullWitness;
            var text = new StringBuilder();
            foreach (string name in new[] { "Type", "Operator", "StopIfTrue", "NumberFormat", "DateOperator", "Text", "TextOperator", "Rank", "Percent", "TopBottom", "AboveBelow", "EqualAverage", "NumStdDev", "DupeUnique", "ScopeType", "PTCondition" }) Witness(text, rule, name);
            object applies = null, areas = null, font = null, borders = null;
            try
            {
                applies = Get(rule, "AppliesTo"); areas = Get(applies, "Areas"); var rectangles = new List<string>();
                var full = new StringBuilder(text.ToString()); var normalized = new StringBuilder(text.ToString());
                foreach (string field in new[] { "Formula1", "Formula2" })
                {
                    string raw = FormulaWitness(field, applies, false), shifted = FormulaWitness(field, applies, true);
                    full.Append(field).Append(':').Append(raw.Length).Append(':').Append(raw).Append(';');
                    normalized.Append(field).Append(':').Append(shifted.Length).Append(':').Append(shifted).Append(';');
                }
                text.Clear();
                for (int i = 1; i <= (int)((dynamic)areas).Count; i++)
                {
                    object area = null, rows = null, columns = null;
                    try
                    {
                        area = ((dynamic)areas).Item[i]; rows = Get(area, "Rows"); columns = Get(area, "Columns");
                        rectangles.Add(((int)((dynamic)area).Row).ToString(CultureInfo.InvariantCulture) + "," + ((int)((dynamic)area).Column).ToString(CultureInfo.InvariantCulture) + "," +
                            ((int)((dynamic)rows).Count).ToString(CultureInfo.InvariantCulture) + "," + ((int)((dynamic)columns).Count).ToString(CultureInfo.InvariantCulture));
                    }
                    finally { ExcelBridge.Release(columns); ExcelBridge.Release(rows); ExcelBridge.Release(area); }
                }
                rectangles.Sort(StringComparer.Ordinal); foreach (string rectangle in rectangles) full.Append("area:").Append(rectangle).Append(';');
                font = Get(rule, "Font");
                foreach (string name in new[] { "Name", "Size", "FontStyle", "ThemeFont", "Bold", "Italic", "Underline", "Strikethrough", "Color", "ColorIndex", "ThemeColor", "TintAndShade" }) Witness(text, font, name);
                borders = Get(rule, "Borders");
                // FormatCondition.Borders uses xlLeft/xlRight/xlTop/xlBottom,
                // not the xlEdge/Inside indexes accepted by Range.Borders.
                foreach (int index in new[] { -4131, -4152, -4160, -4107 })
                {
                    object border = null;
                    try
                    {
                        try { border = ((dynamic)borders).Item[index]; }
                        catch (Exception error) { if (!MissingProperty(error)) throw; text.Append("border:").Append(index).Append(":unavailable;"); continue; }
                        text.Append("border:").Append(index).Append(';');
                        foreach (string name in new[] { "LineStyle", "Weight", "Color", "ColorIndex", "ThemeColor", "TintAndShade" }) Witness(text, border, name);
                    }
                    finally { ExcelBridge.Release(border); }
                }
                full.Append(text); normalized.Append(text);
                using (SHA256 hash = SHA256.Create())
                {
                    fullWitness = Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(full.ToString())));
                    relativeWitness = Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(normalized.ToString())));
                }
                return relative ? relativeWitness : fullWitness;
            }
            finally { ExcelBridge.Release(borders); ExcelBridge.Release(font); ExcelBridge.Release(areas); ExcelBridge.Release(applies); }
        }
        void ReleaseIdentity() { IntPtr held = identity; identity = IntPtr.Zero; if (held != IntPtr.Zero) Marshal.Release(held); }
        public void Dispose()
        { object old = rule; rule = null; try { ExcelBridge.Release(old); } finally { ReleaseIdentity(); GC.SuppressFinalize(this); } }
    }
}
