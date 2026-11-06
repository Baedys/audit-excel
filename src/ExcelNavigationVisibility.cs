using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        const int NavigationSheetVisible = -1;
        static readonly ConditionalWeakTable<object, NavigationVisibilitySession> navigationVisibility =
            new ConditionalWeakTable<object, NavigationVisibilitySession>();

        sealed class NavigationVisibilitySession
        {
            internal readonly Dictionary<long, NavigationVisibilitySheet> Sheets = new Dictionary<long, NavigationVisibilitySheet>();
            internal bool Busy;
            internal void Enter()
            {
                if (Busy) throw new InvalidOperationException("Une modification de visibilité Excel est en cours.");
                Busy = true;
            }
        }

        sealed class NavigationVisibilitySheet
        {
            internal Book Book;
            internal object Worksheet;
            internal int? OriginalVisibility;
            internal readonly HashSet<string> Owners = new HashSet<string>(StringComparer.Ordinal);
            internal readonly List<NavigationVisibilityAxis> Axes = new List<NavigationVisibilityAxis>();
        }

        sealed class NavigationVisibilityAxis
        {
            internal bool Rows;
            internal object Range;
        }

        internal Book CaptureNavigationTargetBook(string fullName)
        {
            var snapshot = new Book { FullName = fullName };
            return WithDirectSnapshotWorkbook(snapshot, workbook => snapshot);
        }

        void SelectNavigationTarget(Book book, object sheet, object range, string owner, bool revealAxes, Action select)
        {
            int previousVisibility = (int)((dynamic)sheet).Visible;
            if (owner == null)
            {
                if (previousVisibility != NavigationSheetVisible)
                    throw new InvalidOperationException("Feuille masquée. Affichez-la dans Excel avant de naviguer.");
                select();
                return;
            }

            NavigationVisibilitySession session = navigationVisibility.GetValue(application, ignored => new NavigationVisibilitySession());
            session.Enter();
            try
            {
                long identity = DirectSession.Identity(sheet);
                NavigationVisibilitySheet state;
                session.Sheets.TryGetValue(identity, out state);
                IReadOnlyList<NavigationAxisRun> hidden = revealAxes ? ReadHiddenNavigationAxes(sheet, range) : new NavigationAxisRun[0];
                bool showSheet = previousVisibility != NavigationSheetVisible;
                if (state == null && (showSheet || hidden.Count != 0))
                    state = new NavigationVisibilitySheet { Book = book, Worksheet = ((dynamic)range).Worksheet };

                var changed = new List<NavigationVisibilityAxis>();
                try
                {
                    if (showSheet)
                    {
                        ((dynamic)sheet).Visible = NavigationSheetVisible;
                        if ((int)((dynamic)sheet).Visible != NavigationSheetVisible)
                            throw new InvalidOperationException("Excel a refusé de dévoiler la feuille.");
                    }
                    if (hidden.Count != 0) UpdateNavigationAxes(sheet, delegate
                    {
                        foreach (NavigationAxisRun run in hidden)
                        {
                            var axis = new NavigationVisibilityAxis { Rows = run.Rows, Range = OpenNavigationAxisRun(sheet, run) };
                            changed.Add(axis);
                            SetNavigationAxisHidden(axis.Range, false);
                        }
                    });
                    select();
                }
                catch (Exception error)
                {
                    var errors = new List<Exception> { error };
                    if (changed.Count != 0)
                    {
                        state.Axes.AddRange(changed);
                        try { RestoreNavigationAxes(sheet, state, changed.AsEnumerable().Reverse().ToArray(), null); }
                        catch (Exception rollback) { errors.Add(rollback); }
                    }
                    if (showSheet)
                    {
                        try { RestoreNavigationSheet(sheet, previousVisibility); }
                        catch (Exception rollback) { state.OriginalVisibility = previousVisibility; errors.Add(rollback); }
                    }
                    if (state != null && !session.Sheets.ContainsKey(identity))
                    {
                        if (state.Axes.Count != 0 || state.OriginalVisibility.HasValue) session.Sheets.Add(identity, state);
                        else Release(state.Worksheet);
                    }
                    if (state != null && session.Sheets.ContainsKey(identity)) state.Owners.Add(owner);
                    if (errors.Count == 1) ExceptionDispatchInfo.Capture(error).Throw();
                    throw new AggregateException("La navigation a échoué et le masquage est à rétablir.", errors);
                }

                if (state == null) return;
                if (showSheet) state.OriginalVisibility = previousVisibility;
                state.Axes.AddRange(changed);
                state.Owners.Add(owner);
                session.Sheets[identity] = state;
            }
            finally { session.Busy = false; }
        }

        internal void RestoreNavigationVisibility(string owner, bool keepSelection)
        { RestoreNavigationVisibilityCore(owner, keepSelection, null); }

        internal void RestoreNavigationVisibilityBeforeSave(object workbook)
        { RestoreNavigationVisibilityCore(null, false, DirectSession.Identity(workbook)); }

        void RestoreNavigationVisibilityCore(string owner, bool keepSelection, long? workbook)
        {
            NavigationVisibilitySession session;
            if (!navigationVisibility.TryGetValue(application, out session)) return;
            if (!session.Busy && session.Sheets.Count == 0) return;
            session.Enter();
            object activeSheet = null, selection = null;
            var errors = new List<Exception>();
            try
            {
                long? activeIdentity = null;
                IReadOnlyList<NavigationAxisRun> keep = null;
                if (keepSelection)
                {
                    activeSheet = ((dynamic)application).ActiveSheet;
                    if (activeSheet != null) activeIdentity = DirectSession.Identity(activeSheet);
                    if (activeIdentity.HasValue && session.Sheets.ContainsKey(activeIdentity.Value))
                    {
                        selection = ((dynamic)application).Selection;
                        if (MapSelectionSupports(selection, MapSelectionRangeInterface)) keep = ReadNavigationAxisProjections(selection);
                    }
                }
                foreach (var pair in session.Sheets.ToArray())
                {
                    NavigationVisibilitySheet state = pair.Value;
                    if (workbook.HasValue && DependencyWorkbookIdentity(state.Book) != workbook.Value) continue;
                    if (owner != null && !state.Owners.Contains(owner)) continue;
                    if (owner != null && state.Owners.Count > 1)
                    {
                        state.Owners.Remove(owner);
                        continue;
                    }
                    try
                    {
                        if (NavigationSheetIsLive(state.Book, pair.Key))
                        {
                            bool preserve = activeIdentity == pair.Key;
                            if (!preserve && state.OriginalVisibility.HasValue)
                                RestoreNavigationSheet(state.Worksheet, state.OriginalVisibility.Value);
                            RestoreNavigationAxes(state.Worksheet, state, state.Axes.ToArray(), preserve ? keep : null);
                        }
                        session.Sheets.Remove(pair.Key);
                        foreach (NavigationVisibilityAxis axis in state.Axes) Release(axis.Range);
                        Release(state.Worksheet);
                    }
                    catch (Exception error) { errors.Add(error); }
                }
            }
            finally { try { Release(selection); Release(activeSheet); } finally { session.Busy = false; } }
            if (errors.Count != 0) throw new AggregateException("Le masquage temporaire n’a pas pu être entièrement rétabli. " +
                String.Join(" ", errors.Select(error => error.Message)), errors);
        }

        bool NavigationSheetIsLive(Book book, long identity)
        {
            if (!DependencyBookIsOpen(book)) return false;
            return WithDirectSnapshotWorkbook(book, delegate(object workbook)
            {
                object sheets = null;
                try
                {
                    sheets = ((dynamic)workbook).Worksheets;
                    for (int index = 1; index <= (int)((dynamic)sheets).Count; index++)
                    {
                        object sheet = null;
                        try { sheet = ((dynamic)sheets).Item[index]; if (DirectSession.Identity(sheet) == identity) return true; }
                        finally { Release(sheet); }
                    }
                    return false;
                }
                finally { Release(sheets); }
            });
        }

        void RestoreNavigationSheet(object sheet, int visibility)
        {
            if ((int)((dynamic)sheet).Visible != NavigationSheetVisible) return;
            object active = null, parent = null, sheets = null;
            try
            {
                active = ((dynamic)application).ActiveSheet;
                if (active != null && DirectSession.Identity(active) == DirectSession.Identity(sheet))
                {
                    // BeforeSave can silently ignore Visible on the active sheet.
                    parent = ((dynamic)sheet).Parent; sheets = ((dynamic)parent).Sheets;
                    for (int index = 1; index <= (int)((dynamic)sheets).Count; index++)
                    {
                        object alternative = null;
                        try
                        {
                            alternative = ((dynamic)sheets).Item[index];
                            if (DirectSession.Identity(alternative) == DirectSession.Identity(sheet) ||
                                (int)((dynamic)alternative).Visible != NavigationSheetVisible) continue;
                            ((dynamic)alternative).Activate();
                            break;
                        }
                        finally { Release(alternative); }
                    }
                }
                OfficeCall(sheet, "Visible", BindingFlags.SetProperty, visibility);
                if ((int)((dynamic)sheet).Visible != visibility)
                    throw new InvalidOperationException("Excel a refusé de remasquer la feuille.");
            }
            finally { Release(sheets); Release(parent); Release(active); }
        }

        static void UpdateNavigationAxes(object sheet, Action action)
        {
            if ((int)((dynamic)sheet).Visible == NavigationSheetVisible) ExcelConditionalFillContext.RunWorksheet(sheet, action);
            else action();
        }

        static void RestoreNavigationAxes(object sheet, NavigationVisibilitySheet state, NavigationVisibilityAxis[] axes, IReadOnlyList<NavigationAxisRun> keep)
        {
            if (axes.Length == 0) return;
            UpdateNavigationAxes(sheet, delegate
            {
                foreach (NavigationVisibilityAxis axis in axes)
                {
                    RestoreNavigationAxis(sheet, axis, keep);
                    state.Axes.Remove(axis); Release(axis.Range);
                }
            });
        }

        static void RestoreNavigationAxis(object sheet, NavigationVisibilityAxis axis, IReadOnlyList<NavigationAxisRun> keep)
        {
            if (ReadNavigationAxisHidden(axis.Range, axis.Rows) == true) return;
            if (keep == null || keep.Count == 0) { HideNavigationAxis(axis.Range, axis.Rows); return; }
            string address = (string)((dynamic)axis.Range).Address[true, true, 1, false];
            Reference rectangle = Reference.Parse("", address);
            if (rectangle == null) throw new InvalidOperationException("Excel ne retrouve plus l’axe dévoilé.");
            int first = axis.Rows ? rectangle.FirstRow : rectangle.FirstColumn;
            int last = axis.Rows ? rectangle.LastRow : rectangle.LastColumn;
            foreach (NavigationAxisRun run in keep.Where(run => run.Rows == axis.Rows && run.Last >= first && run.First <= last).OrderBy(run => run.First))
            {
                if (run.First > first) HideNavigationAxisPart(sheet, axis.Rows, first, run.First - 1);
                first = Math.Max(first, run.Last + 1);
                if (first > last) return;
            }
            HideNavigationAxisPart(sheet, axis.Rows, first, last);
        }

        static void HideNavigationAxisPart(object sheet, bool rows, int first, int last)
        {
            object range = OpenNavigationAxisRun(sheet, new NavigationAxisRun(rows, first, last));
            try { HideNavigationAxis(range, rows); }
            finally { Release(range); }
        }

        static void HideNavigationAxis(object range, bool rows)
        {
            SetNavigationAxisHidden(range, true);
            if (ReadNavigationAxisHidden(range, rows) != true)
                throw new InvalidOperationException("Excel a refusé de remasquer une ligne ou colonne.");
        }
    }
}
