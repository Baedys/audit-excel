using System;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    /// <summary>Owns the two COM references supplied by its caller, not borrowed references.
    /// Rollback preserves the color attributes touched by this port. Merged cells, gradients
    /// and heterogeneous rich-text colors are deliberately unsupported.</summary>
    public sealed class ExcelDirectColorCell : IDirectColorCell
    {
        object range, worksheet;
        readonly long sheetIdentity;
        bool disposed;
        sealed class ColorState
        {
            internal int Index, Rgb;
            internal int? Theme;
            internal double Tint;
            internal bool Same(ColorState other)
            { return Index == other.Index && Rgb == other.Rgb && Theme == other.Theme && Tint == other.Tint; }
        }
        sealed class FormatState
        {
            internal ExcelDirectColorCell Owner;
            internal DirectColorProperty Property;
            internal ColorState Color, PatternColor;
            internal int Pattern;
        }
        public ExcelDirectColorCell(object range, object worksheet)
        {
            if (range == null) throw new ArgumentNullException("range");
            if (worksheet == null) throw new ArgumentNullException("worksheet");
            sheetIdentity = Identity(worksheet);
            this.range = range; this.worksheet = worksheet;
        }
        void Alive() { if (disposed) throw new ObjectDisposedException("ExcelDirectColorCell"); }
        static long Identity(object value)
        {
            IntPtr identity = Marshal.GetIUnknownForObject(value);
            try { return identity.ToInt64(); }
            finally { Marshal.Release(identity); }
        }
        public long SheetIdentity { get { Alive(); return sheetIdentity; } }
        public string CurrentSheetName { get { Alive(); return (string)((dynamic)worksheet).Name; } }
        public string CurrentAddress
        {
            get { Alive(); return (string)((dynamic)range).Address[false, false, 1, false]; }
        }
        public bool IsSingleCell { get { Alive(); return Convert.ToDouble(((dynamic)range).CountLarge) == 1.0; } }
        public void ValidateWrite()
        {
            Alive();
            if (!IsSingleCell) throw new InvalidOperationException("Les couleurs directes exigent une seule cellule.");
            // Merged targets alias several physical cells; a one-cell snapshot cannot restore them.
            if (Convert.ToBoolean(((dynamic)range).MergeCells)) throw new InvalidOperationException("Les cellules fusionnées ne sont pas prises en charge pour les couleurs directes.");
            if (Convert.ToBoolean(((dynamic)worksheet).ProtectContents)) throw new InvalidOperationException("La feuille est protégée.");
            object owner = null;
            try
            {
                owner = ((dynamic)range).Worksheet;
                if (Identity(owner) != SheetIdentity) throw new InvalidOperationException("La cellule ne correspond plus à sa feuille.");
                string address = CurrentAddress;
                if (String.IsNullOrEmpty(address) || address.IndexOf('#') >= 0)
                    throw new InvalidOperationException("La référence de cellule n'est plus valide.");
            }
            finally { Release(owner); }
        }
        static void Property(DirectColorProperty property)
        { if (property != DirectColorProperty.Fill && property != DirectColorProperty.Font) throw new ArgumentOutOfRangeException("property"); }
        object Format(DirectColorProperty property)
        {
            Alive(); Property(property);
            return property == DirectColorProperty.Fill ? ((dynamic)range).Interior : ((dynamic)range).Font;
        }
        static int ScalarInt(object value)
        {
            if (value == null || value == DBNull.Value) throw new InvalidOperationException("Couleurs hétérogènes non prises en charge.");
            return Convert.ToInt32(value);
        }
        static double ScalarDouble(object value)
        {
            if (value == null || value == DBNull.Value) throw new InvalidOperationException("Nuances hétérogènes non prises en charge.");
            return Convert.ToDouble(value);
        }
        static ColorState ReadState(object format, bool pattern)
        {
            dynamic f = format;
            var state = new ColorState {
                Index = ScalarInt(pattern ? f.PatternColorIndex : f.ColorIndex),
                Rgb = ScalarInt(pattern ? f.PatternColor : f.Color),
                Tint = ScalarDouble(pattern ? f.PatternTintAndShade : f.TintAndShade)
            };
            try
            {
                int theme = ScalarInt(pattern ? f.PatternThemeColor : f.ThemeColor);
                if (theme >= 1 && theme <= 12) state.Theme = theme;
            }
            catch (COMException ex)
            {
                // Excel raises 1004 when an explicit RGB/automatic color has no theme binding.
                if ((uint)ex.ErrorCode != 0x800A03EC) throw;
            }
            catch (ArgumentException ex)
            {
                // Font.ThemeColor for an automatic/non-theme font uses E_INVALIDARG.
                if ((uint)ex.HResult != 0x80070057) throw;
            }
            return state;
        }
        public DirectColorValue Read(DirectColorProperty property)
        {
            object format = null;
            try
            {
                format = Format(property); dynamic f = format;
                int index = ScalarInt(f.ColorIndex);
                if (index == -4105) return DirectColorValue.Automatic;
                if (index == -4142) return DirectColorValue.None;
                return DirectColorValue.Rgb(ScalarInt(f.Color));
            }
            finally { Release(format); }
        }
        public object Capture(DirectColorProperty property)
        {
            Property(property); ValidateWrite(); object format = null;
            try
            {
                format = Format(property);
                var state = new FormatState { Owner = this, Property = property, Color = ReadState(format, false) };
                if (property == DirectColorProperty.Fill)
                {
                    state.Pattern = ScalarInt(((dynamic)format).Pattern);
                    if (state.Pattern == 4000 || state.Pattern == 4001)
                        throw new InvalidOperationException("Les dégradés de remplissage ne sont pas pris en charge pour les couleurs directes.");
                    state.PatternColor = ReadState(format, true);
                }
                else ValidateTextColor(state.Color);
                return state;
            }
            finally { Release(format); }
        }
        void ValidateTextColor(ColorState expected)
        {
            // Whole-cell Font.Color can report zero for heterogeneous rich-text runs.
            // Check individual character colors, including their theme/tint bindings.
            if (Convert.ToBoolean(((dynamic)range).HasFormula)) return;
            string text = ((dynamic)range).Value2 as string;
            if (String.IsNullOrEmpty(text)) return;
            for (int i = 1; i <= text.Length; i++)
            {
                object characters = null, font = null;
                try
                {
                    characters = ((dynamic)range).Characters[i, 1]; font = ((dynamic)characters).Font;
                    if (!expected.Same(ReadState(font, false)))
                        throw new InvalidOperationException("Les couleurs de texte enrichi hétérogènes ne sont pas prises en charge.");
                }
                finally { Release(font); Release(characters); }
            }
        }
        public void Write(DirectColorProperty property, DirectColorValue color)
        {
            Alive(); Property(property);
            if (color == null) throw new ArgumentNullException("color");
            if (color.Kind == DirectColorKind.Transparent) return;
            // Also enforce preflight for callers using this port without a transaction.
            Capture(property);
            object format = null;
            try
            {
                format = Format(property); dynamic f = format;
                if (color.Kind == DirectColorKind.Rgb) f.Color = color.Value;
                else f.ColorIndex = color.Kind == DirectColorKind.Automatic ? -4105 : -4142;
            }
            finally { Release(format); }
        }
        static void RestoreState(object format, ColorState state, bool pattern)
        {
            dynamic f = format;
            if (state.Theme.HasValue)
            {
                if (pattern) f.PatternThemeColor = state.Theme.Value; else f.ThemeColor = state.Theme.Value;
            }
            else if (state.Index == -4105 || state.Index == -4142)
            {
                if (pattern) f.PatternColorIndex = state.Index; else f.ColorIndex = state.Index;
            }
            else { if (pattern) f.PatternColor = state.Rgb; else f.Color = state.Rgb; }
            if (pattern) f.PatternTintAndShade = state.Tint; else f.TintAndShade = state.Tint;
        }
        public void Restore(DirectColorProperty property, object snapshot)
        {
            Property(property); Alive();
            var state = snapshot as FormatState;
            if (state == null || !Object.ReferenceEquals(state.Owner, this) || state.Property != property)
                throw new ArgumentException("Instantané de couleur incompatible.", "snapshot");
            ValidateWrite(); object format = null;
            try
            {
                format = Format(property);
                RestoreState(format, state.Color, false);
                if (property == DirectColorProperty.Fill)
                {
                    RestoreState(format, state.PatternColor, true);
                    // Restoring colors can implicitly turn a no-fill cell into a solid fill.
                    ((dynamic)format).Pattern = state.Pattern;
                }
            }
            finally { Release(format); }
        }
        static void Release(object value)
        {
            if (value == null) return;
            try { if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
            catch { /* Disposal must not mask a transaction failure. */ }
        }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true; object oldRange = range, oldSheet = worksheet; range = null; worksheet = null;
            Release(oldRange); Release(oldSheet);
        }
    }
}
