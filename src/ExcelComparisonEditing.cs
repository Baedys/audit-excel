using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal sealed partial class ComparisonCopyReceipt
        {
            internal void SelectEditedRange()
            {
                try
                {
                    Reference target = range == null ? null : bridge.ReadRetainedDependencyRange(book, range);
                    if (target != null) bridge.NavigateDependency(book, new Cell { Sheet = target.Sheet, Address = target.Address });
                }
                catch (Exception error) { Trace.WriteLine("Clair · Sélection après annulation · " + error.Message); }
            }

            internal sealed class EditTarget : IDisposable
            {
                readonly ExcelBridge bridge;
                readonly Book workbook;
                readonly string sheet, address;
                readonly ComparedContent original;
                internal readonly FormulaDisplayStyle Style;
                internal readonly string Text;
                internal readonly bool Mixed;
                object range;

                EditTarget(ExcelBridge bridge, Book workbook, string sheet, string address, object range, ComparedContent original,
                    FormulaDisplayStyle style, string text, bool mixed)
                {
                    this.bridge = bridge; this.workbook = workbook; this.sheet = sheet; this.address = address;
                    this.range = range; this.original = original; Style = style; Text = text; Mixed = mixed;
                }

                internal static EditTarget Open(ExcelBridge bridge, Reference area, FormulaDisplayStyle style, bool mixed)
                {
                    return bridge.AtDirectCell(area.Workbook, area.Sheet, area.Address, (worksheet, range) =>
                    {
                        RequireComparisonCopyTarget(worksheet);
                        ComparedContent snapshot = ReadComparisonCopyTarget(worksheet, range, false);
                        if (snapshot.Merges.Count != 0) throw new InvalidOperationException("La plage d’édition contient des cellules fusionnées.");
                        if (snapshot.Arrays.Count > 0 && (snapshot.Arrays.Count != 1 || snapshot.Arrays[0].Bounds.FirstRow != 1 ||
                            snapshot.Arrays[0].Bounds.FirstColumn != 1 || snapshot.Arrays[0].Bounds.LastRow != snapshot.Rows ||
                            snapshot.Arrays[0].Bounds.LastColumn != snapshot.Columns))
                            throw new InvalidOperationException("Sélectionnez exactement le bloc matriciel entier.");
                        RequireCapturedEditContent(bridge, area, worksheet, snapshot);
                        string text = "";
                        if (!mixed)
                        {
                            object anchor = bridge.RangeAt(worksheet, Coordinates.Address(area.FirstRow, area.FirstColumn));
                            try
                            {
                                string property;
                                text = Convert.ToString(ReadFormulaInStyle(anchor, style, out property), style.Culture);
                            }
                            finally { Release(anchor); }
                        }
                        var lifetime = new Book { FullName = area.Workbook.FullName };
                        directSnapshots.Add(lifetime, RequireDependencySnapshot(area.Workbook));
                        object retained = bridge.RangeAt(worksheet, area.Address);
                        return new EditTarget(bridge, lifetime, area.Sheet, area.Address, retained, snapshot, style, text, mixed);
                    });
                }

                internal ComparisonCopyReceipt Apply(string draft)
                {
                    if (range == null) throw new InvalidOperationException("Le brouillon n’est plus disponible.");
                    ComparisonCopyReceipt receipt = null;
                    try
                    {
                        return bridge.WithComparisonMutation(() => bridge.AtDirectSheet(workbook, sheet, worksheet =>
                        {
                            Reference live = bridge.ReadRetainedDependencyRange(workbook, range);
                            if (live == null || live.Address != address)
                                throw new InvalidOperationException("La plage à modifier a changé dans Excel.");
                            ExcelConditionalFillContext.RunWorksheet(worksheet, () =>
                            {
                                RequireComparisonCopyTarget(worksheet);
                                ComparedContent current = ReadComparisonCopyTarget(worksheet, range, false);
                                if (!SameComparedContent(current, original))
                                    throw new InvalidOperationException("La plage a changé depuis l’ouverture de l’édition.");
                                object retained = bridge.RangeAt(worksheet, address);
                                try
                                {
                                    ComparedContent written;
                                    try
                                    {
                                        WriteDraft(worksheet, draft);
                                        written = ReadComparisonCopyTarget(worksheet, range, false);
                                    }
                                    catch (Exception error)
                                    {
                                        try
                                        {
                                            ComparedContent partial = ReadComparedContent(range);
                                            if (!SameComparedContent(partial, original))
                                                bridge.ReplaceComparedContent(range, partial, original, FormulaDisplayStyle.Invariant, true);
                                        }
                                        catch (Exception rollback) { throw new AggregateException("L’édition et sa restauration Excel ont échoué.", error, rollback); }
                                        throw;
                                    }
                                    if (!SameComparedContent(original, written))
                                    {
                                        receipt = new ComparisonCopyReceipt(bridge, workbook, sheet, retained, original, written);
                                        retained = null;
                                    }
                                }
                                finally { Release(retained); }
                            });
                            return receipt;
                        }));
                    }
                    catch (Exception error)
                    {
                        if (receipt == null) throw;
                        try { receipt.Restore(true); }
                        catch (Exception rollback) { throw new AggregateException("L’édition et son contexte Excel n’ont pas pu être restaurés.", error, rollback); }
                        finally { receipt.Dispose(); }
                        throw;
                    }
                }

                void WriteDraft(object worksheet, string draft)
                {
                    if (original.Arrays.Count != 0)
                    {
                        bridge.WriteLegacyArrayText(range, draft, Style);
                        return;
                    }
                    if (draft.Length == 0) { ((dynamic)range).ClearContents(); return; }
                    if (original.Cells.Length == 1 || Style.R1C1 || !draft.StartsWith("=", StringComparison.Ordinal))
                    {
                        WriteFormula(range, draft, false, Style);
                        return;
                    }
                    Reference bounds = Reference.Parse(sheet, address);
                    object anchor = bridge.RangeAt(worksheet, Coordinates.Address(bounds.FirstRow, bounds.FirstColumn));
                    try
                    {
                        WriteFormula(anchor, draft, false, Style);
                        FormulaDisplayStyle relative = Style.WithNotation(true);
                        string property;
                        string converted = ReadFormulaInStyle(anchor, relative, out property) as string;
                        if (String.IsNullOrEmpty(converted))
                            throw new InvalidOperationException("Excel n’a pas converti la formule du bloc en R1C1.");
                        WriteFormula(range, converted, false, relative);
                    }
                    finally { Release(anchor); }
                }

                public void Dispose()
                {
                    object retained = range; range = null;
                    try { Release(retained); }
                    catch (InvalidComObjectException) { }
                }

                static void RequireCapturedEditContent(ExcelBridge bridge, Reference area, object worksheet, ComparedContent snapshot)
                {
                    for (int index = 0; index < snapshot.Cells.Length; index++)
                    {
                        FormulaCellContent actual = snapshot.Cells[index];
                        if (actual == null) continue;
                        string address = Coordinates.Address(area.FirstRow + index / snapshot.Columns, area.FirstColumn + index % snapshot.Columns);
                        Cell captured = area.Workbook.Get(area.Sheet, address);
                        if (captured.IsFormula ? captured.Formula != actual.Formula : !MatchesCapturedConstant(captured, actual))
                            throw new InvalidOperationException("La plage a changé depuis la comparaison. Comparez de nouveau.");
                    }
                    if (snapshot.Arrays.Count != 0)
                    {
                        string address = Coordinates.Address(area.FirstRow, area.FirstColumn);
                        Cell captured = area.Workbook.Get(area.Sheet, address);
                        if (!captured.IsFormula) throw new InvalidOperationException("Le bloc matriciel a changé depuis la comparaison.");
                        object anchor = bridge.RangeAt(worksheet, address);
                        try { RequireCurrentFormula(anchor, captured); }
                        finally { Release(anchor); }
                    }
                }
            }
        }
    }
}
