using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        static readonly Guid ComparisonChartInterface = new Guid("000208D6-0000-0000-C000-000000000046");

        sealed class ComparisonTab
        {
            internal string Name;
            internal long Identity;
            internal bool IsWorksheet, IsSupported, Visible;
        }

        static ComparisonTab[] ComparisonTabs(object workbook) { return ComparisonTabs(workbook, null); }

        static ComparisonTab[] ComparisonTabs(object workbook, DirectSnapshot snapshot)
        {
            object sheets = ((dynamic)workbook).Sheets;
            try
            {
                var tabs = new ComparisonTab[(int)((dynamic)sheets).Count];
                for (int i = 0; i < tabs.Length; i++)
                {
                    object sheet = ((dynamic)sheets).Item[i + 1];
                    try
                    {
                        bool worksheet = MapSelectionSupports(sheet, MapSelectionWorksheetInterface);
                        bool chart = !worksheet && MapSelectionSupports(sheet, ComparisonChartInterface);
                        IntPtr identity = new IntPtr(DirectSession.Identity(sheet));
                        if (chart && snapshot != null && snapshot.ComparisonCharts.Add(identity)) Marshal.AddRef(identity);
                        tabs[i] = new ComparisonTab { Name = (string)((dynamic)sheet).Name, Identity = identity.ToInt64(),
                            IsWorksheet = worksheet, IsSupported = worksheet || chart,
                            Visible = Convert.ToInt32(((dynamic)sheet).Visible) == -1 };
                    }
                    finally { Release(sheet); }
                }
                return tabs;
            }
            finally { Release(sheets); }
        }

        internal sealed class ComparisonSheetReceipt : IDisposable
        {
            internal readonly Book Source, Destination;
            internal readonly long Identity;
            internal readonly bool IsWorksheet;
            internal object Sheet;

            internal ComparisonSheetReceipt(Book source, Book destination, object sheet, long identity, bool worksheet)
            { Source = source; Destination = destination; Sheet = sheet; Identity = identity; IsWorksheet = worksheet; }

            public void Dispose()
            {
                object sheet = Sheet; Sheet = null;
                try { Release(sheet); }
                catch (InvalidComObjectException) { }
            }
        }

        internal void NavigateComparisonSheet(Reference location)
        {
            WithDirectSnapshotWorkbook(location.Workbook, workbook =>
            {
                object sheets = null, tab = null;
                try
                {
                    sheets = ((dynamic)workbook).Sheets;
                    tab = ((dynamic)sheets).Item[location.Sheet];
                    if (MapSelectionSupports(tab, MapSelectionWorksheetInterface))
                        NavigateDependency(location.Workbook, new Cell { Sheet = location.Sheet, Address = "A1" });
                    else
                    {
                        ((dynamic)workbook).Activate();
                        ((dynamic)tab).Activate();
                    }
                    return true;
                }
                finally { Release(tab); Release(sheets); }
            });
        }

        // Worksheet.Copy activates the new tab. Return to the window hosting the comparison.
        internal void CopyComparisonSheets(Action action)
        {
            WithComparisonMutation(() =>
            {
                dynamic app = application;
                object window = null, sheet = null, selection = null, active = null;
                var group = new List<object>();
                var errors = new List<Exception>();
                bool alerts = app.DisplayAlerts, started = false;
                try
                {
                    window = app.ActiveWindow; sheet = app.ActiveSheet; selection = app.Selection;
                    if (MapSelectionSupports(selection, MapSelectionRangeInterface)) active = app.ActiveCell;
                    if (window != null && sheet != null) ReadDependencySheetGroup(window, group);
                    started = true;
                    app.DisplayAlerts = false;
                    action();
                }
                catch (Exception error) { errors.Add(error); }
                finally
                {
                    if (started)
                    {
                        if (window != null && sheet != null)
                            TryDependencyCleanup(() => RestoreDependencyWindow(window, sheet, selection, active, group), errors);
                        TryDependencyCleanup(() => { app.DisplayAlerts = alerts; }, errors);
                    }
                    foreach (object item in group.Concat(new[] { active, selection, sheet, window }))
                        TryDependencyCleanup(() => Release(item), errors);
                }
                if (errors.Count == 1) ExceptionDispatchInfo.Capture(errors[0]).Throw();
                if (errors.Count > 1) throw new AggregateException("Copie ou restauration du contexte Excel incomplète.", errors);
                return true;
            });
        }

        internal void CopyComparisonSheet(ComparisonRow row, bool toRight, ICollection<ComparisonSheetReceipt> applied)
        {
            Reference source = toRight ? row.Left : row.Right;
            Book destination = row.MissingSheetWorkbook;
            if (source == null || destination == null || DependencyWorkbookIdentity(source.Workbook) == DependencyWorkbookIdentity(destination))
                throw new InvalidOperationException("Sélectionnez un onglet manquant dans l’autre classeur.");
            WithDirectSnapshotWorkbook(source.Workbook, origin => WithDirectSnapshotWorkbook(destination, target =>
            {
                if ((bool)((dynamic)target).ProtectStructure)
                    throw new InvalidOperationException("La structure du classeur de destination est protégée.");
                ComparisonTab[] sourceTabs = ComparisonTabs(origin), targetTabs = ComparisonTabs(target);
                int index = Array.FindIndex(sourceTabs, tab => tab.IsSupported && tab.Name == source.Sheet && tab.Identity == row.MissingSheetIdentity);
                if (index < 0) throw new InvalidOperationException("L’onglet source n’est plus disponible.");
                var targetNames = new HashSet<string>(targetTabs.Select(tab => tab.Name), StringComparer.OrdinalIgnoreCase);
                if (targetNames.Contains(source.Sheet)) throw new InvalidOperationException("Un onglet de même nom existe déjà dans le classeur de destination.");
                string following = sourceTabs.Skip(index + 1).Select(tab => tab.Name).FirstOrDefault(targetNames.Contains);
                string anchorName = following ?? sourceTabs.Take(index).Reverse().Select(tab => tab.Name).FirstOrDefault(targetNames.Contains) ?? targetTabs.Last().Name;
                object from = null, to = null, original = null, anchor = null;
                try
                {
                    from = ((dynamic)origin).Sheets; to = ((dynamic)target).Sheets;
                    original = ((dynamic)from).Item[index + 1]; anchor = ((dynamic)to).Item[anchorName];
                    WithComparisonConditionalBases(origin, sourceTabs[index], () =>
                    {
                        Exception copyFailure = null;
                        try
                        {
                            if (following == null) ((dynamic)original).Copy(Type.Missing, anchor);
                            else ((dynamic)original).Copy(anchor, Type.Missing);
                        }
                        catch (COMException error) { copyFailure = error; }
                        var additions = ComparisonTabs(target).Where(tab => !targetNames.Contains(tab.Name)).ToArray();
                        foreach (ComparisonTab tab in additions)
                        {
                            object copied = ((dynamic)to).Item[tab.Name];
                            applied.Add(new ComparisonSheetReceipt(source.Workbook, destination, copied, DirectSession.Identity(copied), tab.IsWorksheet));
                        }
                        if (copyFailure != null) ExceptionDispatchInfo.Capture(copyFailure).Throw();
                        if (additions.Length != 1 || additions[0].Name != source.Sheet)
                            throw new InvalidOperationException("Excel n’a pas créé l’onglet attendu.");
                    });
                    return true;
                }
                finally { Release(anchor); Release(original); Release(to); Release(from); }
            }));
        }

        void WithComparisonConditionalBases(object workbook, ComparisonTab sheet, Action copy)
        {
            if (!sheet.IsWorksheet) { copy(); return; }
            DirectSession direct = DirectColors;
            direct.Enter();
            ConditionalBook book = null;
            bool sourceWasSaved = false;
            var originals = new List<Tuple<ConditionalEntry, ConditionalFillState>>();
            var errors = new List<Exception>();
            try
            {
                ConditionalSession session = ConditionalFills;
                session.Prune();
                book = session.Book(workbook);
                AdoptConditionalFills(workbook, session, book);
                var groups = book.Entries.Where(entry => entry.SheetIdentity == sheet.Identity).GroupBy(entry => entry.ManagedRuleKey).ToList();
                if (book.MetadataRepair || groups.Any(group => group.Any(entry => entry.Pending)) ||
                    book.Unresolved.Any(entry => String.Equals(entry.Sheet, sheet.Name, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Restaurez les fonds conditionnels en attente sur la feuille source avant de la copier.");
                if (groups.Count != 0) sourceWasSaved = (bool)((dynamic)workbook).Saved;
                foreach (var group in groups)
                {
                    ConditionalEntry entry = group.First();
                    WithConditionalEntryPort(entry, port =>
                    {
                        ConditionalFillState current = port.Read();
                        if (group.Any(member => !ConditionalStateKnown(member, current)))
                            throw new InvalidOperationException("Un fond conditionnel de la feuille source a changé. La copie est interrompue.");
                        if (!entry.Base.Equals(current))
                        {
                            originals.Add(Tuple.Create(entry, current));
                            port.Restore(entry.Base);
                        }
                        return true;
                    });
                }
                copy();
            }
            catch (Exception error) { errors.Add(error); }
            finally
            {
                for (int i = originals.Count - 1; i >= 0; i--)
                {
                    var original = originals[i];
                    try
                    {
                        WithConditionalEntryPort(original.Item1, port =>
                        {
                            if (!port.Read().Equals(original.Item2))
                            {
                                port.Mask();
                                if (!port.Read().Equals(original.Item2)) port.Restore(original.Item2);
                            }
                            return true;
                        });
                    }
                    catch (Exception error)
                    {
                        original.Item1.Pending = true;
                        errors.Add(error);
                        TryDependencyCleanup(() => WithConditionalEntryPort(original.Item1, port =>
                        {
                            ConditionalFillState observed = port.Read();
                            if (observed.Equals(original.Item2)) original.Item1.Pending = false;
                            else RememberOwnConditionalState(book, original.Item1, observed);
                            return true;
                        }), errors);
                    }
                }
                if (sourceWasSaved && originals.Count != 0 && errors.Count == 0)
                    TryDependencyCleanup(() => { ((dynamic)workbook).Saved = true; }, errors);
                direct.Exit();
            }
            if (errors.Count == 1) ExceptionDispatchInfo.Capture(errors[0]).Throw();
            if (errors.Count > 1) throw new AggregateException("La copie ou la restauration des fonds conditionnels est incomplète.", errors);
        }

        internal void ValidateComparisonSheetUndo(IList<ComparisonSheetReceipt> receipts)
        {
            foreach (var batch in receipts.GroupBy(item => DependencyWorkbookIdentity(item.Destination)))
                WithDirectSnapshotWorkbook(batch.First().Destination, workbook =>
                {
                    if ((bool)((dynamic)workbook).ProtectStructure)
                        throw new InvalidOperationException("La structure du classeur de destination est protégée.");
                    ComparisonTab[] tabs = ComparisonTabs(workbook);
                    var names = new HashSet<string>(tabs.Select(tab => tab.Name), StringComparer.OrdinalIgnoreCase);
                    var copied = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    object sheets = ((dynamic)workbook).Sheets;
                    try
                    {
                        foreach (ComparisonSheetReceipt receipt in batch)
                        {
                            object current = null;
                            try
                            {
                                string name = receipt.Sheet == null ? null : (string)((dynamic)receipt.Sheet).Name;
                                if (name != null && names.Contains(name)) current = ((dynamic)sheets).Item[name];
                                if (current == null || DirectSession.Identity(current) != receipt.Identity)
                                    throw new InvalidOperationException("Un onglet copié a été supprimé ou déplacé dans un autre classeur.");
                                copied.Add(name);
                            }
                            catch (COMException error)
                            { throw new InvalidOperationException("Un onglet copié a été supprimé ou déplacé dans un autre classeur.", error); }
                            finally { Release(current); }
                        }
                    }
                    finally { Release(sheets); }
                    if (!tabs.Any(tab => tab.Visible && !copied.Contains(tab.Name)))
                        throw new InvalidOperationException("Le classeur doit conserver au moins un onglet visible.");
                    return true;
                });
        }

        internal void UndoComparisonSheets(List<ComparisonSheetReceipt> receipts)
        {
            ValidateComparisonSheetUndo(receipts);
            WithComparisonMutation(() =>
            {
                dynamic app = application;
                bool alerts = app.DisplayAlerts;
                try
                {
                    app.DisplayAlerts = false;
                    while (receipts.Count != 0)
                    {
                        ComparisonSheetReceipt receipt = receipts[0];
                        ((dynamic)receipt.Sheet).Delete();
                        receipts.RemoveAt(0);
                        receipt.Dispose();
                    }
                    return true;
                }
                finally { app.DisplayAlerts = alerts; }
            });
        }

        internal void LocalizeComparisonSheet(ComparisonSheetReceipt receipt)
        {
            if (!receipt.IsWorksheet) return;
            var worksheets = WithDirectSnapshotWorkbook(receipt.Destination, ComparisonTabs)
                .Where(tab => tab.IsWorksheet).ToDictionary(tab => tab.Name, tab => tab.Name, StringComparer.OrdinalIgnoreCase);
            object used = ((dynamic)receipt.Sheet).UsedRange;
            try
            {
                var bounds = MapSelectionRectangle(used);
                var addresses = SpecialCellAddresses((dynamic)used, -4123, bounds.Top, bounds.Left, bounds.Height, bounds.Width);
                foreach (string address in addresses)
                {
                    object cell = null, block = null, parent = null;
                    try
                    {
                        cell = RangeAt(receipt.Sheet, address);
                        if (!(bool)((dynamic)cell).HasFormula) continue;
                        string formula;
                        if (IsLegacyArrayMember(cell))
                        {
                            block = ((dynamic)cell).CurrentArray;
                            if ((int)((dynamic)block).Row != (int)((dynamic)cell).Row || (int)((dynamic)block).Column != (int)((dynamic)cell).Column) continue;
                            formula = ReadLegacyArrayFormula(cell);
                            if (IsLegacyDataTable(formula)) continue;
                        }
                        else
                        {
                            if (LegacyArrayHasSpill(cell))
                            {
                                parent = ((dynamic)cell).SpillParent;
                                if ((string)((dynamic)parent).Address != (string)((dynamic)cell).Address) continue;
                            }
                            formula = ReadFormula(cell) as string;
                        }
                        string replacement = LocalizeComparisonFormula(formula, receipt.Source, worksheets);
                        if (formula == replacement) continue;
                        if (block != null) RestoreLegacyArrayFormula(block, replacement);
                        else
                        {
                            string format = FormulaNumberFormat(cell);
                            try { WriteFormula(cell, replacement); }
                            finally { RestoreFormulaNumberFormat(cell, format); }
                        }
                    }
                    finally { Release(parent); Release(block); Release(cell); }
                }
            }
            finally { Release(used); }
        }

        internal static string LocalizeComparisonFormula(string formula, Book source, IDictionary<string, string> worksheets)
        {
            if (String.IsNullOrEmpty(formula) || formula.IndexOf('[') < 0) return formula;
            IReadOnlyList<FormulaSyntax.Token> tokens;
            try { tokens = FormulaSyntax.PresentationTokens(formula, FormulaDisplayStyle.Invariant, null); }
            catch (FormatException) { return formula; }
            StringBuilder rewritten = null;
            int previous = 0;
            for (int i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (token.Kind != "Atom") continue;
                var qualifier = FormulaParser.ReadExternalSheet(formula, token.Start);
                string name;
                if (qualifier == null || !String.Equals(qualifier.Workbook, source.Name, StringComparison.OrdinalIgnoreCase) ||
                    !worksheets.TryGetValue(qualifier.Sheet, out name) || !FormulaParser.AddressPattern.IsMatch(formula, qualifier.AddressStart)) continue;
                if ((i >= 2 && tokens[i - 1].Text == ":" && tokens[i - 2].Kind == "Atom" && tokens[i - 2].Text.Contains("!")) ||
                    (i + 2 < tokens.Count && tokens[i + 1].Text == ":" && tokens[i + 2].Kind == "Atom" && tokens[i + 2].Text.Contains("!"))) continue;
                if (qualifier.Path.Length != 0)
                {
                    string path = NormalizeDependencyWorkbookPath(qualifier.Path + qualifier.Workbook);
                    if (path == null || !MatchesDependencyWorkbookPath(path, source.FullName)) continue;
                }
                if (rewritten == null) rewritten = new StringBuilder(formula.Length);
                rewritten.Append(formula, previous, token.Start - previous).Append('\'').Append(name.Replace("'", "''")).Append("'!");
                previous = qualifier.AddressStart;
            }
            return rewritten == null ? formula : rewritten.Append(formula, previous, formula.Length - previous).ToString();
        }
    }
}
