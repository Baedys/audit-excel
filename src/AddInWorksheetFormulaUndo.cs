using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace FormulaExplorer
{
    public sealed partial class ExcelAddIn
    {
        static readonly Guid FormulaUndoEvents = new Guid("00024413-0000-0000-C000-000000000046");
        const int SheetChangeEvent = 1564;
        WorksheetFormulaUndo worksheetFormulaUndo;
        Action<object, object> worksheetFormulaChanged;

        public bool IsWorksheetFormulaUndoLoaded() { return worksheetFormulaUndo != null; }

        void EnsureWorksheetFormulaUndo()
        {
            if (worksheetFormulaUndo != null) return;
            var session = new WorksheetFormulaUndo(new ExcelBridge(application), worksheetShortcut.RegisterFormulaUndo, worksheetShortcut.ClearFormulaUndoIfCurrent);
            Action<object, object> changed = (sheet, target) => session.SheetChanged();
            try { ComEventsHelper.Combine(application, FormulaUndoEvents, SheetChangeEvent, changed); }
            catch { session.Dispose(); throw; }
            worksheetFormulaChanged = changed;
            worksheetFormulaUndo = session;
        }

        void DisconnectWorksheetFormulaUndo()
        {
            WorksheetFormulaUndo session = worksheetFormulaUndo;
            Action<object, object> changed = worksheetFormulaChanged;
            worksheetFormulaUndo = null;
            worksheetFormulaChanged = null;
            if (changed != null && application != null)
                try { ComEventsHelper.Remove(application, FormulaUndoEvents, SheetChangeEvent, changed); }
                catch (Exception error) { Trace.WriteLine("Clair · Événements d’annulation Excel · " + error.Message); }
            if (session != null)
                try { session.Dispose(); }
                catch (Exception error) { Trace.WriteLine("Clair · Annulation Excel · " + error.Message); }
        }

        void DispatchWorksheetFormulaUndo()
        { if (worksheetFormulaUndo != null) worksheetFormulaUndo.Run(); }
    }
}
