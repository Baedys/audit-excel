using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace FormulaExplorer
{
    public sealed partial class ExcelBridge
    {
        internal Book RefreshDependencyWorkbook(Book snapshot)
        {
            RequireDependencySnapshot(snapshot);
            return WithDirectSnapshotWorkbook(snapshot, Capture);
        }

        internal long DependencyWorkbookIdentity(Book snapshot)
        { return RequireDependencySnapshot(snapshot).Workbook.ToInt64(); }

        internal Dictionary<string, string> DependencyWorksheetNames(Book previous, Book current)
        {
            DirectSnapshot before = RequireDependencySnapshot(previous), after = RequireDependencySnapshot(current);
            var names = after.Sheets.ToDictionary(sheet => sheet.Value, sheet => sheet.Key);
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sheet in before.Sheets)
            {
                string name;
                if (names.TryGetValue(sheet.Value, out name)) result.Add(sheet.Key, name);
            }
            return result;
        }

        internal object RetainDependencyRange(Book book, Reference range)
        { return AtDirectSheet(book, range.Sheet, sheet => RangeAt(sheet, range.Address)); }

        internal Reference ReadRetainedDependencyRange(Book book, object range)
        {
            object sheet = null;
            try
            {
                sheet = ((dynamic)range).Worksheet;
                string name = (string)((dynamic)sheet).Name;
                IntPtr identity;
                if (!RequireDependencySnapshot(book).Sheets.TryGetValue(name, out identity) ||
                    identity.ToInt64() != DirectSession.Identity(sheet)) return null;
                Reference result = Reference.Parse(name, (string)((dynamic)range).Address);
                if (result != null) result.Workbook = book;
                return result;
            }
            catch (COMException) { return null; }
            finally { Release(sheet); }
        }

        internal static string NormalizeDependencyWorkbookPath(string path)
        {
            Uri address;
            if (!Uri.TryCreate(path, UriKind.Absolute, out address)) return null;
            if (address.IsFile)
            {
                if (address.Query.Length != 0 || address.Fragment.Length != 0) return null;
                try
                {
                    string local = address.LocalPath;
                    string root = Path.GetPathRoot(local);
                    if (String.IsNullOrEmpty(root) || root.Length < 3 || String.IsNullOrEmpty(Path.GetFileName(local))) return null;
                    return Path.GetFullPath(local);
                }
                catch (ArgumentException) { return null; }
                catch (NotSupportedException) { return null; }
                catch (PathTooLongException) { return null; }
            }
            if ((address.Scheme != Uri.UriSchemeHttps && address.Scheme != Uri.UriSchemeHttp) ||
                String.IsNullOrEmpty(address.Host) || address.UserInfo.Length != 0 || address.Fragment.Length != 0) return null;
            string name = Uri.UnescapeDataString(address.Segments.Last());
            return name.Length == 0 || name.IndexOfAny(new[] { '/', '\\' }) >= 0 ? null : address.AbsoluteUri;
        }

        internal void OpenDependencyWorkbook(Book owner, string fullName)
        {
            RequireDependencySnapshot(owner);
            string path = NormalizeDependencyWorkbookPath(fullName);
            if (path == null) throw new ArgumentException("Le lien exige le chemin complet d’un classeur.", "fullName");
            string name = IsDependencyWorkbookUrl(path) ? Uri.UnescapeDataString(new Uri(path).Segments.Last()) : Path.GetFileName(path);
            WithDirectSnapshotWorkbook(owner, delegate(object workbook)
            {
                object books = null;
                try
                {
                    books = ((dynamic)application).Workbooks;
                    bool homonym = false;
                    var priorBooks = new HashSet<long>();
                    int count = (int)((dynamic)books).Count;
                    for (int index = 1; index <= count; index++)
                    {
                        object candidate = null;
                        try
                        {
                            candidate = ((dynamic)books).Item[index];
                            if (MatchesDependencyWorkbookPath(path, (string)((dynamic)candidate).FullName)) return true;
                            homonym |= String.Equals(name, (string)((dynamic)candidate).Name, StringComparison.OrdinalIgnoreCase);
                            priorBooks.Add(DirectSession.Identity(candidate));
                        }
                        finally { Release(candidate); }
                    }
                    if (homonym) throw new InvalidOperationException("Un classeur nommé « " + name + " » est déjà ouvert depuis un autre emplacement.");
                    OpenDependencyFile(books, path, priorBooks);
                    return true;
                }
                finally { Release(books); }
            });
        }

        static bool IsDependencyWorkbookUrl(string path)
        { return path.StartsWith("http://", StringComparison.Ordinal) || path.StartsWith("https://", StringComparison.Ordinal); }

        internal static bool MatchesDependencyWorkbookPath(string normalizedPath, string actual)
        {
            return String.Equals(normalizedPath, NormalizeDependencyWorkbookPath(actual),
                IsDependencyWorkbookUrl(normalizedPath) ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
        }

        void OpenDependencyFile(object books, string path, HashSet<long> priorBooks)
        {
            dynamic app = application;
            bool events = app.EnableEvents, screen = app.ScreenUpdating;
            int security = (int)app.AutomationSecurity;
            object priorWindow = null, priorSheet = null, priorSelection = null, priorActive = null, opened = null;
            var priorGroup = new List<object>();
            bool restoreSettings = false, attemptedOpen = false;
            var errors = new List<Exception>();
            try
            {
                priorWindow = app.ActiveWindow; priorSheet = app.ActiveSheet; priorSelection = app.Selection;
                if (priorWindow == null || priorSheet == null)
                    throw new InvalidOperationException("Activez une fenêtre et une feuille Excel avant d’ouvrir le lien.");
                if (MapSelectionSupports(priorSelection, MapSelectionRangeInterface)) priorActive = app.ActiveCell;
                ReadDependencySheetGroup(priorWindow, priorGroup);
                restoreSettings = true;
                app.EnableEvents = false; app.ScreenUpdating = false;
                app.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
                attemptedOpen = true;
                opened = ((dynamic)books).Open(path, 0);
                if (opened == null || !MatchesDependencyWorkbookPath(path, (string)((dynamic)opened).FullName))
                    throw new InvalidOperationException("Excel n’a pas ouvert le classeur au chemin demandé.");
            }
            catch (Exception error)
            {
                errors.Add(error);
                if (opened != null) TryDependencyCleanup(delegate
                {
                    if (!priorBooks.Contains(DirectSession.Identity(opened))) ((dynamic)opened).Close(false);
                }, errors);
            }
            finally
            {
                if (attemptedOpen) TryDependencyCleanup(delegate { RestoreDependencyWindow(priorWindow, priorSheet, priorSelection, priorActive, priorGroup); }, errors);
                if (restoreSettings)
                {
                    TryDependencyCleanup(delegate { app.AutomationSecurity = security; }, errors);
                    TryDependencyCleanup(delegate { app.ScreenUpdating = screen; }, errors);
                    TryDependencyCleanup(delegate { app.EnableEvents = events; }, errors);
                }
                DirectSession colorSession;
                if (errors.Count == 0 && events && opened != null && directSessions.TryGetValue(application, out colorSession))
                {
                    ExcelDirectColorEvents listener = colorSession.ColorEvents;
                    if (listener != null) TryDependencyCleanup(delegate { listener.WorkbookOpen(opened); }, errors);
                }
                foreach (object item in new[] { opened, priorActive, priorSelection, priorSheet, priorWindow })
                    TryDependencyCleanup(delegate { Release(item); }, errors);
                foreach (object sheet in priorGroup) TryDependencyCleanup(delegate { Release(sheet); }, errors);
            }
            if (errors.Count == 1) ExceptionDispatchInfo.Capture(errors[0]).Throw();
            if (errors.Count > 1) throw new AggregateException("Ouverture du lien ou restauration du contexte Excel incomplète.", errors);
        }

        void RestoreDependencyWindow(object window, object sheet, object selection, object activeCell, List<object> group)
        {
            dynamic app = application;
            object currentSheet = null, currentSelection = null, currentActive = null;
            try
            {
                ((dynamic)window).Activate();
                if (!DependencyWindowIsActive(window))
                {
                    // A hidden Excel instance can defer activation while redraw is suppressed.
                    app.ScreenUpdating = true;
                    ((dynamic)window).Activate();
                    if (!DependencyWindowIsActive(window))
                        throw new InvalidOperationException("La fenêtre Excel initiale n’a pas pu être rétablie.");
                }
                bool restoreGroup = !DependencySheetGroupMatches(window, group);
                if (restoreGroup) ((dynamic)sheet).Select(true);
                currentSheet = app.ActiveSheet;
                if (currentSheet == null || DirectSession.Identity(currentSheet) != DirectSession.Identity(sheet))
                {
                    ((dynamic)sheet).Activate();
                    Release(currentSheet); currentSheet = null;
                    currentSheet = app.ActiveSheet;
                    if (currentSheet == null || DirectSession.Identity(currentSheet) != DirectSession.Identity(sheet))
                        throw new InvalidOperationException("La feuille Excel initiale n’a pas pu être rétablie.");
                }
                if (selection != null)
                {
                    currentSelection = app.Selection;
                    if (!SameDependencySelection(currentSelection, selection))
                    {
                        ((dynamic)selection).Select();
                        if (MapSelectionSupports(selection, MapSelectionRangeInterface))
                        {
                            Release(currentSelection); currentSelection = null;
                            currentSelection = app.Selection;
                            if (!SameDependencySelection(currentSelection, selection))
                                throw new InvalidOperationException("La sélection Excel initiale n’a pas pu être rétablie.");
                        }
                    }
                }
                if (activeCell != null)
                {
                    currentActive = app.ActiveCell;
                    if (!SameDependencySelection(currentActive, activeCell))
                    {
                        ((dynamic)activeCell).Activate();
                        Release(currentActive); currentActive = null;
                        currentActive = app.ActiveCell;
                        if (!SameDependencySelection(currentActive, activeCell))
                            throw new InvalidOperationException("La cellule active initiale n’a pas pu être rétablie.");
                    }
                }
                if (restoreGroup)
                {
                    long sheetId = DirectSession.Identity(sheet);
                    foreach (object grouped in group)
                        if (DirectSession.Identity(grouped) != sheetId) ((dynamic)grouped).Select(false);
                    ((dynamic)sheet).Activate();
                    if (!DependencySheetGroupMatches(window, group))
                        throw new InvalidOperationException("Le groupe de feuilles Excel initial n’a pas pu être rétabli.");
                }
            }
            finally { Release(currentActive); Release(currentSelection); Release(currentSheet); }
        }

        bool DependencyWindowIsActive(object window)
        {
            object current = null;
            try
            {
                current = ((dynamic)application).ActiveWindow;
                return current != null && (int)((dynamic)current).Hwnd == (int)((dynamic)window).Hwnd;
            }
            finally { Release(current); }
        }

        static void ReadDependencySheetGroup(object window, List<object> group)
        {
            object sheets = null;
            try
            {
                sheets = ((dynamic)window).SelectedSheets;
                int count = (int)((dynamic)sheets).Count;
                for (int index = 1; index <= count; index++) group.Add(((dynamic)sheets).Item[index]);
            }
            finally { Release(sheets); }
        }

        static bool DependencySheetGroupMatches(object window, List<object> group)
        {
            var actual = new List<object>();
            try
            {
                ReadDependencySheetGroup(window, actual);
                return actual.Count == group.Count && new HashSet<long>(actual.Select(DirectSession.Identity)).SetEquals(group.Select(DirectSession.Identity));
            }
            finally { foreach (object sheet in actual) Release(sheet); }
        }

        static bool SameDependencySelection(object current, object prior)
        {
            if (current == null) return false;
            if (Marshal.IsComObject(current) && Marshal.IsComObject(prior) && DirectSession.Identity(current) == DirectSession.Identity(prior)) return true;
            return MapSelectionSupports(current, MapSelectionRangeInterface) && MapSelectionSupports(prior, MapSelectionRangeInterface) &&
                String.Equals((string)((dynamic)current).Address, (string)((dynamic)prior).Address, StringComparison.OrdinalIgnoreCase);
        }

        static void TryDependencyCleanup(Action cleanup, ICollection<Exception> errors)
        {
            try { cleanup(); }
            catch (Exception error) { errors.Add(error); }
        }

        internal IReadOnlyList<Book> CaptureDependencyPeers(Book primary)
        {
            RequireDependencySnapshot(primary);
            return WithDirectSnapshotWorkbook(primary, delegate(object workbook)
            {
                long primaryIdentity = DirectSession.Identity(workbook);
                var peers = new List<Book>();
                object books = null;
                try
                {
                    books = ((dynamic)application).Workbooks;
                    int count = (int)((dynamic)books).Count;
                    for (int index = 1; index <= count; index++)
                    {
                        object candidate = null;
                        string name = "n° " + index;
                        try
                        {
                            candidate = ((dynamic)books).Item[index];
                            if (DirectSession.Identity(candidate) == primaryIdentity) continue;
                            name = (string)((dynamic)candidate).Name;
                            peers.Add(Capture(candidate));
                        }
                        catch (Exception error)
                        {
                            if (!(error is COMException) && !(error is InvalidOperationException) &&
                                !(error is InvalidComObjectException) && !(error is RuntimeBinderException)) throw;
                            string warning = "Classeur « " + name + " » ignoré pendant la capture des liens. " + error.Message;
                            if (!primary.Warnings.Contains(warning)) primary.Warnings.Add(warning);
                        }
                        finally { Release(candidate); }
                    }
                }
                finally { Release(books); }
                return peers.AsReadOnly();
            });
        }

        internal void NavigateDependency(Book book, Cell cell, string revealOwner = null, bool revealAxes = false)
        {
            RequireDependencySnapshot(book);
            if (cell == null) throw new ArgumentNullException("cell");
            if (String.IsNullOrWhiteSpace(cell.Sheet) || String.IsNullOrWhiteSpace(cell.Address))
                throw new ArgumentException("La navigation exige une adresse locale de feuille.", "cell");
            Reference[] areas = cell.Address.Split(',').Select(area => Reference.Parse(cell.Sheet, area.Trim())).ToArray();
            if (areas.Any(area => area == null))
                throw new ArgumentException("La navigation exige une adresse locale de feuille.", "cell");
            string address = String.Join(",", areas.Select(area => area.Address));
            AtDirectCell(book, cell.Sheet, address, delegate(object sheet, object range)
            {
                SelectNavigationTarget(book, sheet, range, revealOwner, revealAxes, delegate { ((dynamic)application).Goto(range, true); });
                return true;
            });
        }

        internal bool DependencyBookIsOpen(Book snapshot)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            DirectSnapshot ticket;
            if (!directSnapshots.TryGetValue(snapshot, out ticket)) return false;
            long identity = ticket.Workbook.ToInt64();
            object books = null;
            try
            {
                books = ((dynamic)application).Workbooks;
                int count = (int)((dynamic)books).Count;
                for (int index = 1; index <= count; index++)
                {
                    object workbook = null;
                    try
                    {
                        workbook = ((dynamic)books).Item[index];
                        if (DirectSession.Identity(workbook) == identity) return true;
                    }
                    finally { Release(workbook); }
                }
                return false;
            }
            finally { Release(books); }
        }

        static DirectSnapshot RequireDependencySnapshot(Book book)
        {
            if (book == null) throw new ArgumentNullException("book");
            DirectSnapshot ticket;
            if (!directSnapshots.TryGetValue(book, out ticket))
                throw new InvalidOperationException("Cette analyse exige une capture physique du classeur Excel.");
            return ticket;
        }
    }
}
