using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FormulaExplorer
{
    public sealed partial class MainForm
    {
        sealed class ExplorerExpressionState
        {
            internal string Current, Anchor;
            internal string[] Selected, Expanded;
            internal bool OnGrid;
        }

        sealed class ExplorerLiveSource : IDisposable
        {
            readonly object[] ranges;
            readonly Book[] books;

            internal ExplorerLiveSource(ExcelBridge bridge, Book primary, Cell root, TraceLocation source)
            {
                Reference[] targets = new[] { TraceCellReference(root) }
                    .Concat(source == null ? Enumerable.Empty<Reference>() : source.Ranges).ToArray();
                ranges = new object[targets.Length];
                books = new Book[targets.Length];
                try
                {
                    for (int index = 0; index < targets.Length; index++)
                    {
                        books[index] = targets[index].Workbook ?? primary;
                        ranges[index] = bridge.RetainDependencyRange(books[index], targets[index]);
                    }
                }
                catch { Dispose(); throw; }
            }

            internal Reference[] Read(ExcelBridge bridge, Func<Book, Book> replace)
            {
                var result = new Reference[ranges.Length];
                for (int index = 0; index < result.Length; index++)
                {
                    Book current = replace(books[index]);
                    if (current != null) result[index] = bridge.ReadRetainedDependencyRange(current, ranges[index]);
                    if (result[index] == null && (index != 0 || ranges.Length == 1)) return null;
                }
                return result;
            }

            internal void Rebind(Func<Book, Book> replace)
            { for (int index = 0; index < books.Length; index++) books[index] = replace(books[index]); }

            public void Dispose()
            {
                for (int index = 0; index < ranges.Length; index++)
                {
                    ExcelBridge.Release(ranges[index]);
                    ranges[index] = null;
                }
            }
        }

        ExplorerLiveSource liveSource;
        int explorerWindow;

        void ClearExplorerHistory()
        {
            foreach (ExplorerHistory item in history)
            {
                if (item.LiveSource != null) item.LiveSource.Dispose();
                if (item.Navigation != null) item.Navigation.Dispose();
            }
            history.Clear();
        }

        void RefreshExplorerModel(TreeView navigationView)
        { RefreshExplorerModelCore(navigationView, false); }

        void RefreshExplorerModelCore(TreeView navigationView, bool fromWorksheetHistory)
        {
            if (bridge == null || book == null || liveSource == null)
                throw new InvalidOperationException("Chargez une sélection Excel pour actualiser le tracé.");
            if (!fromWorksheetHistory && navigationView != null && IsDocked && explorerWindow != 0 &&
                explorerWindow != bridge.ActiveNavigationWindowHandle()) return;

            Book previous = book;
            Book fresh = bridge.RefreshDependencyWorkbook(previous);
            IReadOnlyList<Book> peers = bridge.CaptureDependencyPeers(fresh);
            var oldBooks = traceEngine.TraceWorkbooks.ToDictionary(bridge.DependencyWorkbookIdentity);
            var replacements = new Dictionary<Book, Book>();
            var sheetsByBook = new Dictionary<Book, Dictionary<string, string>>();
            var sheetKeys = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Book next in new[] { fresh }.Concat(peers))
            {
                Book old;
                if (!oldBooks.TryGetValue(bridge.DependencyWorkbookIdentity(next), out old)) continue;
                next.TraceId = old.TraceId;
                replacements.Add(old, next);
                var names = bridge.DependencyWorksheetNames(old, next);
                sheetsByBook.Add(old, names);
                foreach (var name in names) sheetKeys.Add(old.TraceId + name.Key, next.TraceId + name.Value);
            }
            Func<Book, Book> replace = old =>
            {
                Book next;
                return replacements.TryGetValue(old ?? previous, out next) ? next : null;
            };
            Func<Book, string, string> sheetName = (old, name) =>
            {
                Dictionary<string, string> names;
                string next;
                return sheetsByBook.TryGetValue(old ?? previous, out names) && names.TryGetValue(name, out next) ? next : null;
            };

            Reference[] sources = liveSource.Read(bridge, replace);
            if (sources == null)
                throw new InvalidOperationException("Les cellules sources du tracé ne sont plus accessibles.");
            Cell nextRoot = TraceSourceCell(sources[0] ?? sources[1]);
            TraceLocation nextSource = RefreshTraceSource(traceSource, sources, replace, sheetName);
            ExplorerSelection selection = RefreshRelationPaths(CaptureExplorerSelection(), root, nextRoot, sheetKeys);
            Cell oldExpression = expressionOwner;
            string expressionSheet = oldExpression == null || !oldExpression.IsFormula ? null :
                sheetName(oldExpression.Workbook, oldExpression.Sheet);
            Cell nextExpression = expressionSheet == null ? null : ReferenceEquals(oldExpression, root) ? nextRoot :
                replace(oldExpression.Workbook).Get(expressionSheet, oldExpression.Address);
            ExplorerExpressionState expressionState = navigationView == null || nextExpression == null ? null : CaptureExplorerExpressions();

            Cell nextCaptured = ReferenceEquals(capturedRoot, root) ? nextRoot : null;
            var previousHistory = history.Reverse().ToArray();
            var refreshedHistory = new List<ExplorerHistory>();
            foreach (ExplorerHistory item in previousHistory)
            {
                Reference[] saved = item.LiveSource == null ? null : item.LiveSource.Read(bridge, replace);
                if (saved == null) continue;
                Cell refreshed = TraceSourceCell(saved[0] ?? saved[1]);
                if (ReferenceEquals(item.Cell, capturedRoot)) nextCaptured = refreshed;
                refreshedHistory.Add(new ExplorerHistory {
                    Cell = refreshed, LiveSource = item.LiveSource, Reverse = item.Reverse, HighlightOwner = item.HighlightOwner,
                    Expressions = item.Expressions, Navigation = item.Navigation, ExpressionNavigation = item.ExpressionNavigation,
                    Source = RefreshTraceSource(item.Source, saved, replace, sheetName),
                    Selection = RefreshRelationPaths(item.Selection, item.Cell, refreshed, sheetKeys)
                });
            }
            string primarySheet = nextCaptured == null ? sheetName(previous, capturedRoot.Sheet) ?? fresh.SheetOrder.First() : nextCaptured.Sheet;
            UpdateCapturedBook(fresh, primarySheet, nextCaptured == null ? capturedRoot.Address : nextCaptured.Address, peers);
            root = nextRoot;
            traceSource = nextSource;
            liveSource.Rebind(replace);
            history.Clear();
            var retained = new HashSet<ExplorerLiveSource>(refreshedHistory.Select(item => item.LiveSource));
            foreach (ExplorerHistory item in previousHistory)
                if (item.LiveSource == null || !retained.Contains(item.LiveSource))
                {
                    if (item.LiveSource != null) item.LiveSource.Dispose();
                    if (item.Navigation != null) item.Navigation.Dispose();
                    RemoveTraceHighlightOwner(item.HighlightOwner);
                }
            foreach (ExplorerHistory item in refreshedHistory)
            {
                item.LiveSource.Rebind(replace);
                history.Push(item);
            }

            ResetExplorerNavigation();
            ResetFlowDisplay();
            status.Text = "Tracé actualisé · " + TraceCellLabel(root);
            ShowTreeSelection(selection);
            if (fromWorksheetHistory && bridge.ActiveNavigationWindowHandle() != explorerWindow)
                bridge.ActivateExplorerWindow(explorerWindow);
            bool reselect = navigationView != null && IsDocked && explorerWindow != 0 &&
                explorerWindow == bridge.ActiveNavigationWindowHandle();
            if (expressionState != null && expressionOwner != null && expressionOwner.Workbook == nextExpression.Workbook &&
                String.Equals(expressionOwner.Key, nextExpression.Key, StringComparison.OrdinalIgnoreCase))
                RestoreExplorerExpressions(expressionState, reselect && navigationView == expressions);
            if (reselect && navigationView == tree) FollowRelationSelection(true);
        }

        static Cell TraceSourceCell(Reference source)
        { return source.Workbook.Get(source.Sheet, Coordinates.Address(source.FirstRow, source.FirstColumn)); }

        static TraceLocation RefreshTraceSource(TraceLocation old, Reference[] sources, Func<Book, Book> replace,
            Func<Book, string, string> sheetName)
        {
            if (old == null) return null;
            var ranges = sources.Skip(1).ToArray();
            string launch = old.SeparateSources ? sheetName(old.LaunchWorkbook, old.LaunchSheet) : null;
            return new TraceLocation(ranges, old.SeparateSources ? launch ?? ranges[0].Sheet : null,
                launch == null ? ranges[0].Workbook : replace(old.LaunchWorkbook));
        }

        static ExplorerSelection RefreshRelationPaths(ExplorerSelection state, Cell oldRoot, Cell newRoot,
            Dictionary<string, string> sheets)
        {
            string oldRootKey = new WorkbookSheet(oldRoot.Workbook, oldRoot.Sheet).CellKey(oldRoot.Address);
            string newRootKey = new WorkbookSheet(newRoot.Workbook, newRoot.Sheet).CellKey(newRoot.Address);
            Func<string[], string[]> remap = path => path == null ? null : path.Select((key, index) =>
            {
                if (index == 0 && String.Equals(key, oldRootKey, StringComparison.OrdinalIgnoreCase)) return newRootKey;
                int separator = key.IndexOf('\0');
                if (separator < 0) return key;
                int prefix = key.LastIndexOf(':', separator) + 1;
                string sheet;
                return sheets.TryGetValue(key.Substring(prefix, separator - prefix), out sheet) ?
                    key.Substring(0, prefix) + sheet + key.Substring(separator) : "";
            }).ToArray();
            return new ExplorerSelection {
                Current = remap(state.Current), Anchor = remap(state.Anchor),
                Selected = state.Selected.Select(remap).ToArray(), Expanded = state.Expanded.Select(remap).ToArray()
            };
        }

        static string ExpressionPath(Dictionary<TreeNode, string> paths, TreeNode node)
        {
            string path;
            return node != null && paths.TryGetValue(node, out path) ? path : null;
        }

        Dictionary<TreeNode, string> ExpressionPaths()
        {
            var result = new Dictionary<TreeNode, string>();
            AddExpressionPaths(expressions.Nodes, "", result);
            return result;
        }

        static void AddExpressionPaths(TreeNodeCollection nodes, string prefix, Dictionary<TreeNode, string> paths)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (TreeNode node in nodes)
            {
                var group = node.Tag as CriterionGroup;
                var expression = node.Tag as ExpressionNode;
                string key;
                if (group != null) key = "Critère " + group.Range.Text + "\0" + group.Criterion.Text;
                else if (expression.Kind == "Function" || expression.Kind == "Invocation")
                    key = expression.Kind + " " + expression.Text.Split('(')[0].TrimEnd().ToUpperInvariant();
                else if (expression.Children.Count != 0 && (expression.Kind == "Group" || expression.Kind.StartsWith("Operator ", StringComparison.Ordinal)))
                    key = expression.Kind;
                else key = expression.Kind + " " + expression.Text;
                int occurrence;
                counts.TryGetValue(key, out occurrence);
                counts[key] = occurrence + 1;
                string path = prefix + "/" + occurrence + ":" + key.Length + ":" + key;
                paths.Add(node, path);
                AddExpressionPaths(node.Nodes, path, paths);
            }
        }

        ExplorerExpressionState CaptureExplorerExpressions()
        {
            if (expressions.Nodes.Count == 0) return null;
            var paths = ExpressionPaths();
            return new ExplorerExpressionState {
                Current = ExpressionPath(paths, expressions.SelectedNode), Anchor = ExpressionPath(paths, expressionAnchor),
                Selected = selectedExpressions.Select(node => ExpressionPath(paths, node)).ToArray(),
                Expanded = paths.Where(pair => pair.Key.IsExpanded).Select(pair => pair.Value).ToArray(), OnGrid = expressionGridActive
            };
        }

        void RestoreExplorerExpressions(ExplorerExpressionState state, bool follow)
        {
            if (state == null || expressions.Nodes.Count == 0) return;
            var paths = ExpressionPaths().ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.Ordinal);
            Func<string, TreeNode> find = path =>
            {
                TreeNode node;
                return path != null && paths.TryGetValue(path, out node) ? node : null;
            };
            bool wasBuilding = buildingExpressions;
            buildingExpressions = true;
            expressions.BeginUpdate();
            try
            {
                expressions.Nodes[0].Collapse();
                foreach (TreeNode node in state.Expanded.Select(find).Where(node => node != null)) node.Expand();
                selectedExpressions.Clear();
                selectedExpressions.AddRange(state.Selected.Select(find).Where(node => node != null));
                expressions.SelectedNode = find(state.Current) ?? selectedExpressions.LastOrDefault() ?? expressions.Nodes[0];
                if (selectedExpressions.Count == 0 && state.Selected.Length != 0) selectedExpressions.Add(expressions.SelectedNode);
                expressionAnchor = find(state.Anchor) ?? expressions.SelectedNode;
                expressions.SelectedNode.EnsureVisible();
            }
            finally { expressions.EndUpdate(); buildingExpressions = wasBuilding; }
            RefreshExpressionSelection(follow);
            if (!follow && state.OnGrid) UpdateExpressionGridSelection(false);
        }
    }
}
