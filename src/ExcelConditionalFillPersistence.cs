using System;
using System.Collections.Generic;
using System.Linq;

namespace FormulaExplorer
{
    public sealed class SavedHighlightRestoreResult
    {
        public DirectHighlightRestoreResult Direct { get; internal set; }
        public ConditionalFillRestoreResult Conditional { get; internal set; }
        public System.Collections.ObjectModel.ReadOnlyCollection<string> Errors { get; internal set; }
    }
    public sealed partial class ExcelBridge
    {
        sealed class ExcelConditionalXmlStore : IDirectHighlightXmlStore, IDisposable
        {
            object parts;
            readonly List<ExcelDirectXmlPart> retained = new List<ExcelDirectXmlPart>();
            readonly Action<string, string> mutation;
            public ExcelConditionalXmlStore(object workbook, Action<string, string> mutation = null) { parts = ((dynamic)workbook).CustomXMLParts; this.mutation = mutation; }
            public IEnumerable<IDirectHighlightXmlPart> OwnParts()
            {
                object selected = null; var result = new List<IDirectHighlightXmlPart>();
                try
                {
                    selected = ((dynamic)parts).SelectByNamespace(ConditionalFillDocument.Namespace);
                    if ((int)((dynamic)selected).Count > 67) throw new FormatException("Trop de parties XML de fonds conditionnels.");
                    for (int i = 1; i <= (int)((dynamic)selected).Count; i++)
                    {
                        object part = null;
                        try { part = ((dynamic)selected).Item[i]; var item = new ExcelDirectXmlPart(part, mutation, "Conditional"); retained.Add(item); result.Add(item); part = null; }
                        finally { Release(part); }
                    }
                    return result;
                }
                finally { Release(selected); }
            }
            public IDirectHighlightXmlPart Add(string xml)
            {
                object part = null;
                try
                {
                    if (mutation != null) mutation("BeforeConditionalXmlAdd", "");
                    part = ((dynamic)parts).Add(xml);
                    if (mutation != null) mutation("AfterConditionalXmlAdd", (string)((dynamic)part).Id);
                    var item = new ExcelDirectXmlPart(part, mutation, "Conditional"); retained.Add(item); part = null; return item;
                }
                finally { Release(part); }
            }
            public void Dispose() { foreach (ExcelDirectXmlPart part in retained) part.Dispose(); retained.Clear(); Release(parts); parts = null; }
        }
        static ConditionalFillDocument ReadConditionalDocuments(IEnumerable<IDirectHighlightXmlPart> parts)
        {
            long characters = 0; var documents = new List<ConditionalFillDocument>(); var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (IDirectHighlightXmlPart part in parts)
            {
                if (part == null || String.IsNullOrEmpty(part.Id) || !ids.Add(part.Id) || ids.Count > 64) throw new FormatException("Parties XML CF invalides ou trop nombreuses.");
                string xml = part.Xml;
                if (xml == null || (characters += xml.Length) > ConditionalFillDocument.MaxXmlChars) throw new FormatException("Journal des fonds conditionnels trop volumineux.");
                documents.Add(ConditionalFillDocument.FromXml(xml));
            }
            return ConditionalFillDocument.Merge(documents);
        }
        static string ConditionalSignature(string xml) { return ConditionalFillDocument.FromXml(xml).ToXml(); }
        static bool KnownConditionalPart(ConditionalBook book, IDirectHighlightXmlPart part)
        {
            string signature = ConditionalSignature(part.Xml), known;
            if (book.KnownParts.TryGetValue(part.Id, out known) && known == signature) return true;
            // Compensation can recreate an identical journal under another
            // Office part ID. Contents, rather than a recycled ID, prove it known.
            if (!book.KnownParts.Values.Contains(signature)) return false;
            book.KnownParts[part.Id] = signature; return true;
        }
        static void RememberConditionalParts(ConditionalBook book, IEnumerable<IDirectHighlightXmlPart> parts)
        { foreach (IDirectHighlightXmlPart part in parts) book.KnownParts[part.Id] = ConditionalSignature(part.Xml); }
        public ConditionalFillDocument ReadSavedConditionalFillData(object workbook)
        {
            DirectSession direct = DirectColors; direct.Enter();
            try { CheckDirectWorkbook(workbook); using (var store = new ExcelConditionalXmlStore(workbook)) return ReadConditionalDocuments(store.OwnParts()); }
            finally { direct.Exit(); }
        }
        public ConditionalFillDocument ReadUnloadedConditionalFillData(object workbook)
        {
            DirectSession direct = DirectColors; direct.Enter();
            try
            {
                CheckDirectWorkbook(workbook); ConditionalSession session = ConditionalFills; session.Prune(); ConditionalBook book = session.Book(workbook);
                using (var store = new ExcelConditionalXmlStore(workbook)) return ReadConditionalDocuments(store.OwnParts().Where(part => !KnownConditionalPart(book, part)));
            }
            finally { direct.Exit(); }
        }
        public void SaveHighlightData(object workbook)
        {
            DirectSession direct = DirectColors; direct.Enter();
            try { SaveHighlightDataCore(workbook, direct, false); }
            finally { try { direct.ReleaseUnusedLeases(); } finally { direct.Exit(); } }
        }
        void SaveHighlightDataCore(object workbook, DirectSession direct, bool allowPending)
        {
            CheckDirectWorkbook(workbook); direct.Reconcile(allowPending);
            if (direct.Transaction.PendingRepair.Count != 0) throw new InvalidOperationException("Les couleurs directes incomplètement restaurées empêchent l'enregistrement.");
            ConditionalSession session = ConditionalFills; session.Prune(); ConditionalBook book = session.Book(workbook);
            AdoptConditionalFills(workbook, session, book);
            if (!allowPending) session.CheckRepair();
            using (var directStore = new ExcelDirectXmlStore(workbook, session.Step))
            using (var conditionalStore = new ExcelConditionalXmlStore(workbook, session.Step))
            {
                DirectHighlightDocument directDocument = PrepareDirectHighlightSave(workbook, direct, directStore);
                ConditionalFillDocument conditionalDocument = ProjectConditionalFills(workbook, session, book);
                try { direct.Update(delegate { CombinedHighlightXmlStorage.Replace(directStore, directDocument, conditionalStore, conditionalDocument); }); }
                catch (Exception error)
                {
                    if (error.Data.Contains("InitialStateRestored") && Object.Equals(error.Data["InitialStateRestored"], false))
                    { book.MetadataRepair = true; book.RepairReceipt = error.Data["RepairReceipt"] as CombinedHighlightXmlRepairReceipt; }
                    throw;
                }
                direct.RememberReplacement(workbook, directDocument, directStore.LastAddedId);
                book.KnownParts.Clear(); RememberConditionalParts(book, conditionalStore.OwnParts());
                book.MetadataRepair = false;
                book.RepairReceipt = null;
            }
        }
        public void RepairSavedHighlightData(object workbook)
        {
            ConditionalSession session;
            if (!conditionalSessions.TryGetValue(DirectHighlightSessionId, out session)) return;
            DirectSession direct = DirectColors; direct.Enter();
            try
            {
                CheckDirectWorkbook(workbook); session.Prune(); ConditionalBook book;
                if (session.Books.TryGetValue(DirectSession.Identity(workbook), out book)) RepairSavedHighlightDataCore(workbook, direct, book);
            }
            finally { direct.Exit(); }
        }
        void RepairSavedHighlightDataCore(object workbook, DirectSession direct, ConditionalBook book)
        {
            if (!book.MetadataRepair) return;
            if (book.RepairReceipt == null) throw new InvalidOperationException("Le reçu de réparation des journaux de couleurs est indisponible. Les parties existantes sont conservées.");
            ConditionalSession session = ConditionalFills;
            using (var directStore = new ExcelDirectXmlStore(workbook, session.Step))
            using (var conditionalStore = new ExcelConditionalXmlStore(workbook, session.Step))
            {
                CombinedHighlightXmlRepairReceipt receipt = book.RepairReceipt;
                direct.Update(delegate { CombinedHighlightXmlStorage.Repair(directStore, conditionalStore, receipt); });
                var parts = directStore.OwnParts().ToList();
                direct.RememberReplacement(workbook, receipt.DirectDocument, receipt.DirectDocument.Count == 0 ? null : parts.Single().Id);
                book.KnownParts.Clear(); RememberConditionalParts(book, conditionalStore.OwnParts());
                book.MetadataRepair = false; book.RepairReceipt = null;
            }
        }
        public SavedHighlightRestoreResult RestoreSavedHighlightData(object workbook, DirectHighlightRestoreChoice choice)
        {
            ConditionalFillRestoreResult conditional = RestoreConditionalFillData(workbook, choice);
            // Keeping direct colors never authorizes forgetting a CF base.
            if (choice == DirectHighlightRestoreChoice.KeepPermanent && conditional.RemainingCount != 0)
                return new SavedHighlightRestoreResult { Conditional = conditional, Direct = new DirectHighlightRestoreResult(), Errors = conditional.Errors };
            DirectHighlightRestoreResult direct = RestoreDirectHighlightData(workbook, choice);
            return new SavedHighlightRestoreResult { Direct = direct, Conditional = conditional, Errors = conditional.Errors.Concat(direct.Errors).ToList().AsReadOnly() };
        }
    }
}
