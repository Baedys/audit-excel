using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using FormulaExplorer;

// STA UI callbacks against an in-memory fixture; never connects to Excel or saves preferences.
public static class FlowTreePanelTests
{
    static int checks;
    const BindingFlags Private = BindingFlags.Instance | BindingFlags.NonPublic;
    static T Field<T>(MainForm form, string name) { return (T)typeof(MainForm).GetField(name, Private).GetValue(form); }
    static void Set(MainForm form, string name, object value) { typeof(MainForm).GetField(name, Private).SetValue(form, value); }
    static void Invoke(MainForm form, string name, params object[] args)
    {
        try { typeof(MainForm).GetMethod(name, Private).Invoke(form, args); Application.DoEvents(); }
        catch (TargetInvocationException ex) { throw ex.InnerException ?? ex; }
    }
    static void Equal<T>(T actual, T expected, string context)
    {
        checks++;
        if (!Object.Equals(actual, expected)) throw new Exception("Flow tree UI " + context + ": expected " + expected + ", got " + actual);
    }
    static IEnumerable<TreeNode> Walk(TreeNodeCollection roots)
    {
        foreach (TreeNode node in roots) { yield return node; foreach (TreeNode child in Walk(node.Nodes)) yield return child; }
    }
    static FlowTreeNode Model(TreeNode node) { return (FlowTreeNode)node.Tag; }
    static TreeNode Node(MainForm form, FlowTreeNodeKind kind, FlowCategory? category, string sheet, string address)
    {
        return Walk(Field<TreeView>(form, "flowGroups").Nodes).Single(n => {
            var model = n.Tag as FlowTreeNode;
            return model != null && model.Kind == kind && model.Category == category && model.Sheet == sheet &&
                (address == null || model.Block != null && model.Block.Address == address);
        });
    }
    static TreeNode Root(MainForm form) { return Field<TreeView>(form, "flowGroups").Nodes[0]; }
    static void State(TreeNode node, CheckState expected, string context) { Equal(node.StateImageIndex, (int)expected, context); }
    static DataGridViewRow Row(MainForm form, FlowCategory category, string sheet, string address)
    {
        return Field<DataGridView>(form, "flow").Rows.Cast<DataGridViewRow>().Single(r => {
            var block = (FlowBlock)r.Tag; return block.Category == category && block.Sheet == sheet && block.Address == address;
        });
    }
    static void Key(Control control, Keys key)
    {
        var args = new KeyEventArgs(key);
        typeof(Control).GetMethod("OnKeyDown", Private).Invoke(control, new object[] { args });
        Application.DoEvents();
        Equal(args.Handled, true, key + " handled");
        Equal(args.SuppressKeyPress, true, key + " suppresses default key behavior");
    }
    static Book Fixture()
    {
        var book = new Book { Name = "Flow tree UI fixture", FullName = "in-memory-flow-tree" };
        book.SetSheetIndex("Z", 1); book.SetSheetIndex("A", 2);
        Demo.Add(book, "Z", "B1", "1", null, "Nombre");
        Demo.Add(book, "Z", "B2", "2", null, "Nombre");
        Demo.Add(book, "Z", "C1", "3", "=SUM(B1:B2)", "Nombre");
        Demo.Add(book, "Z", "D1", "6", "='A'!B1", "Nombre");
        Demo.Add(book, "A", "B1", "6", "='Z'!C1+B2", "Nombre");
        Demo.Add(book, "A", "B2", "3", null, "Nombre");
        return book;
    }
    public static int Run()
    {
        checks = 0;
        using (var form = new MainForm(null))
        {
            Set(form, "flowViewSettingsFile", null); Set(form, "flowViewSettings", new FlowViewSettings());
            Set(form, "flowPalette", FlowPalette.Default());
            var group = Field<CheckBox>(form, "flowGroupBySheet"); group.Checked = false;
            form.Show(); form.ExecuteRibbonCommand("flow");
            var limitedBook = new Book { Name = "Hidden sheet budget" };
            Demo.Add(limitedBook, "Only", "A1", "1", null, "Nombre");
            Demo.Add(limitedBook, "Only", "A3", "2", null, "Nombre");
            Demo.Add(limitedBook, "Only", "B1", "3", "=SUM(A1,A3)", "Nombre");
            var limitedResult = FlowAnalysis.Analyze(new DependencyEngine(limitedBook), limitedBook.Cells.Values, limitedBook.Cells.Values);
            var hiddenSheet = FlowTree.Build(limitedResult, FlowBlocks.Build(limitedResult), false, false, null).Children[0].Children[0];
            Equal(hiddenSheet.IsHidden, true, "budget fixture has hidden sheet");
            Equal(hiddenSheet.Blocks.Count, 2, "budget fixture has direct block children");
            var limitedTree = Field<TreeView>(form, "flowGroups"); limitedTree.Nodes.Clear(); Set(form, "flowTreeTruncated", false);
            object[] budgetArgs = { hiddenSheet, limitedTree.Nodes, 1 };
            typeof(MainForm).GetMethod("AddFlowTreeNode", Private).Invoke(form, budgetArgs);
            Equal((int)budgetArgs[2], 0, "hidden sheet exhausts exact render budget");
            Equal(Walk(limitedTree.Nodes).Count(), 1, "hidden sheet renders only one budgeted node");
            Equal(Field<bool>(form, "flowTreeTruncated"), true, "hidden sheet marks omitted descendants");
            Invoke(form, "SetBook", Fixture(), "A", "B1"); Invoke(form, "AnalyzeFlow");
            TreeView tree = Field<TreeView>(form, "flowGroups"); DataGridView grid = Field<DataGridView>(form, "flow");
            var view = Field<ComboBox>(form, "flowView"); var details = Field<CheckBox>(form, "flowCellDetails");
            Equal(Model(Root(form)).RoleCount, 6, "fixture role count");
            Equal(grid.Rows.Count, 5, "fixture block count");
            State(Root(form), CheckState.Unchecked, "fresh root");
            Equal(Model(Root(form).Nodes[0]).Kind, FlowTreeNodeKind.Category, "ungrouped category first");
            group.Checked = true;
            Equal(Model(Root(form).Nodes[0]).Sheet, "Z", "worksheet order first");
            Equal(Model(Root(form).Nodes[1]).Sheet, "A", "worksheet order second");
            Invoke(form, "ToggleFlowTreeNode", Node(form, FlowTreeNodeKind.Sheet, null, "Z", null));
            State(Node(form, FlowTreeNodeKind.Sheet, null, "Z", null), CheckState.Checked, "target sheet enabled");
            State(Node(form, FlowTreeNodeKind.Sheet, null, "A", null), CheckState.Unchecked, "other sheet untouched");
            State(Root(form), CheckState.Indeterminate, "root mixed across sheets");
            Equal(Row(form, FlowCategory.Inputs, "A", "B2").Cells[0].Value, (object)CheckState.Unchecked, "other-sheet grid unchanged");

            TreeNode input = Node(form, FlowTreeNodeKind.Block, FlowCategory.Inputs, "Z", "B1:B2");
            tree.SelectedNode = input; Application.DoEvents(); string selectedKey = Model(input).Key;
            Equal(((FlowBlock)grid.CurrentRow.Tag).Address, "B1:B2", "tree selects matching grid block");
            group.Checked = false;
            Equal(Model(tree.SelectedNode).Key, selectedKey, "block selection survives grouping");
            State(Node(form, FlowTreeNodeKind.Category, FlowCategory.Inputs, null, null), CheckState.Indeterminate, "global inputs mixed");
            State(Node(form, FlowTreeNodeKind.Category, FlowCategory.Calculations, null, null), CheckState.Indeterminate, "global calculations mixed");
            group.Checked = true;
            Equal(Model(tree.SelectedNode).Key, selectedKey, "block selection survives reverse grouping");
            details.Checked = true;
            Equal(grid.Rows.Count, 6, "detail rows");
            grid.CurrentCell = Row(form, FlowCategory.Inputs, "Z", "B2").Cells[0]; Application.DoEvents();
            Equal(Model(tree.SelectedNode).Key, selectedKey, "detail second cell selects containing tree block");
            Key(grid, Keys.Space);
            Equal(((FlowBlock)grid.CurrentRow.Tag).Address, "B2", "detail toggle does not jump to B1");
            Equal(grid.CurrentCell.ColumnIndex, 0, "detail toggle keeps checkbox column");
            Equal(Row(form, FlowCategory.Inputs, "Z", "B1").Cells[0].Value, (object)CheckState.Checked, "first input remains active");
            Equal(Row(form, FlowCategory.Inputs, "Z", "B2").Cells[0].Value, (object)CheckState.Unchecked, "second input disabled");
            State(Node(form, FlowTreeNodeKind.Block, FlowCategory.Inputs, "Z", "B1:B2"), CheckState.Indeterminate, "tree block mixed");
            State(Node(form, FlowTreeNodeKind.Category, FlowCategory.Inputs, "Z", null), CheckState.Indeterminate, "sheet inputs mixed");
            State(Node(form, FlowTreeNodeKind.Sheet, null, "Z", null), CheckState.Indeterminate, "sheet mixed after detail toggle");
            tree.SelectedNode = Node(form, FlowTreeNodeKind.Block, FlowCategory.Inputs, "Z", "B1:B2");
            Key(tree, Keys.Space);
            State(Node(form, FlowTreeNodeKind.Block, FlowCategory.Inputs, "Z", "B1:B2"), CheckState.Checked, "Space enables entire mixed block");
            Equal(Row(form, FlowCategory.Inputs, "Z", "B2").Cells[0].Value, (object)CheckState.Checked, "tree toggle updates detail grid");
            details.Checked = false;

            view.SelectedIndex = 1;
            Equal(group.Checked, false, "boundary setting initially independent");
            group.Checked = true; view.SelectedIndex = 0;
            Equal(group.Checked, true, "input grouping retained after boundary change");
            Equal(Model(tree.SelectedNode).Key, selectedKey, "input selection retained across views");
            group.Checked = false; view.SelectedIndex = 1;
            Equal(group.Checked, true, "boundary grouping retained independently");
            Equal(Field<FlowViewSettings>(form, "flowViewSettings").InputsBySheet, false, "stored input option");
            Equal(Field<FlowViewSettings>(form, "flowViewSettings").FlowsBySheet, true, "stored boundary option");
            Field<ComboBox>(form, "flowFocus").SelectedIndex = 1; Invoke(form, "AnalyzeFlow");
            Equal(Model(Root(form)).RoleCount, 4, "four boundary roles including double role");
            Equal(grid.Rows.Count, 4, "double-role cell has two rows");
            TreeNode inside = Node(form, FlowTreeNodeKind.Category, FlowCategory.InsideInflows, null, null);
            Equal(Model(inside.Nodes[0]).Kind, FlowTreeNodeKind.Block, "focus sheet wrapper hidden");
            Equal(Model(inside.Nodes[0]).Sheet, "A", "focus block belongs to root sheet despite workbook order");
            TreeNode outside = Node(form, FlowTreeNodeKind.Category, FlowCategory.OutsidePrecedents, null, null);
            Equal(Model(outside.Nodes[0]).Kind, FlowTreeNodeKind.Sheet, "nonfocus sheet wrapper visible");
            Equal(Model(outside.Nodes[0]).Sheet, "Z", "nonfocus sheet retained");
            Invoke(form, "ToggleFlowTreeNode", inside);
            Equal(Row(form, FlowCategory.InsideInflows, "A", "B1").Cells[0].Value, (object)CheckState.Checked, "inflow enabled");
            Equal(Row(form, FlowCategory.InsideOutflows, "A", "B1").Cells[0].Value, (object)CheckState.Unchecked, "same-cell outflow independent");
            State(Root(form), CheckState.Indeterminate, "boundary root mixed");
            Invoke(form, "ToggleFlowTreeNode", Node(form, FlowTreeNodeKind.Category, FlowCategory.InsideOutflows, null, null));
            Equal(Row(form, FlowCategory.InsideInflows, "A", "B1").Cells[0].Value, (object)CheckState.Checked, "outflow activation preserves inflow");
            Equal(Row(form, FlowCategory.InsideOutflows, "A", "B1").Cells[0].Value, (object)CheckState.Checked, "outflow enabled separately");
            tree.SelectedNode = Node(form, FlowTreeNodeKind.Block, FlowCategory.OutsidePrecedents, "Z", "C1");
            string boundaryKey = Model(tree.SelectedNode).Key;
            group.Checked = false;
            Equal(Model(tree.SelectedNode).Key, boundaryKey, "boundary block survives grouping");
            Equal(tree.SelectedNode.Text.StartsWith("'Z'!C1", StringComparison.Ordinal), true, "ungrouped external block identifies sheet");
            view.SelectedIndex = 0; view.SelectedIndex = 1;
            Equal(Model(tree.SelectedNode).Key, boundaryKey, "boundary selection survives view round-trip");
            tree.SelectedNode = Root(form); Key(tree, Keys.Space);
            State(Root(form), CheckState.Checked, "Space activates all boundary roles");
            Equal(grid.Rows.Cast<DataGridViewRow>().All(r => Object.Equals(r.Cells[0].Value, CheckState.Checked)), true, "root toggle updates all grid roles");
            tree.SelectedNode = Node(form, FlowTreeNodeKind.Block, FlowCategory.OutsidePrecedents, "Z", "C1");
            TreeNode stale = tree.SelectedNode;
            Key(tree, Keys.Enter);
            Equal(Field<Cell>(form, "root").Key, Book.Key("Z", "C1"), "Enter explores block first cell");
            Equal(Field<TabControl>(form, "tabs").SelectedIndex, 0, "Enter shows explorer");
            Equal(tree.Nodes.Count, 0, "exploration clears old tree");
            Equal(grid.Rows.Count, 0, "exploration clears old grid");
            Invoke(form, "ToggleFlowTreeNode", stale); Invoke(form, "ExploreFlowTreeNode", stale);
            Equal(Field<FlowAnalysisResult>(form, "flowResult"), null, "stale nodes cannot revive analysis");
            Equal(Field<List<FlowActivation>>(form, "flowActivations").Count, 0, "stale nodes cannot activate colors");
            Invoke(form, "Back");
            Equal(Field<Cell>(form, "root").Key, Book.Key("A", "B1"), "Back returns previous root");
            Invoke(form, "ToggleFlowTreeNode", stale); Invoke(form, "ExploreFlowTreeNode", stale);
            Equal(Field<Cell>(form, "root").Key, Book.Key("A", "B1"), "stale exploration after Back ignored");
            Equal(tree.Nodes.Count, 0, "Back leaves no obsolete tree");
            Equal(Field<List<FlowActivation>>(form, "flowActivations").Count, 0, "Back leaves no obsolete activations");
            form.Close();
        }
        return checks;
    }
}
