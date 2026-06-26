using LocalDocAI.Models;
using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace LocalDocAI.WordIntegration
{
    public class DocumentContextManager
    {
        private readonly Word.Application _app;

        public DocumentContextManager(Word.Application app)
        {
            _app = app;
        }

        public DocumentContext GetContext(bool includeComments = false, bool includeRevisions = false,
            bool includeHeadings = false, bool includeTables = false)
        {
            var ctx = new DocumentContext();
            var doc = GetActiveDoc();
            if (doc == null) return ctx;

            try
            {
                ctx.DocumentFileName = doc.Name;
                ctx.SelectedText = GetSelectedText();
                ctx.FullText = GetFullText(doc);

                if (includeHeadings) ctx.Headings = GetHeadings(doc);
                if (includeComments) ctx.Comments = CommentReader.ReadAll(doc);
                if (includeRevisions) ctx.Revisions = RevisionReader.ReadAll(doc);
                if (includeTables) ctx.TableSummaries = GetTableSummaries(doc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetContext error: " + ex.Message);
            }

            return ctx;
        }

        public string GetSelectedText()
        {
            try
            {
                var sel = _app.Selection;
                if (sel?.Type == Word.WdSelectionType.wdSelectionNormal)
                    return sel.Range.Text ?? "";
                return "";
            }
            catch { return ""; }
        }

        public string GetFullText(Word.Document doc = null)
        {
            try
            {
                doc = doc ?? GetActiveDoc();
                if (doc == null) return "";
                return doc.Content.Text ?? "";
            }
            catch { return ""; }
        }

        public List<string> GetHeadings(Word.Document doc = null)
        {
            var result = new List<string>();
            try
            {
                doc = doc ?? GetActiveDoc();
                if (doc == null) return result;

                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    var styleObj = p.get_Style() as Word.Style;
                    if (styleObj == null) continue;
                    var name = styleObj.NameLocal ?? "";
                    if (name.StartsWith("Heading") || name.StartsWith("Tiêu đề"))
                        result.Add(p.Range.Text?.TrimEnd('\r') ?? "");
                }
            }
            catch { }
            return result;
        }

        public List<string> GetTableSummaries(Word.Document doc = null)
        {
            var result = new List<string>();
            try
            {
                doc = doc ?? GetActiveDoc();
                if (doc == null) return result;
                int tableNum = 0;

                foreach (Word.Table table in doc.Tables)
                {
                    tableNum++;
                    var sb = new StringBuilder();
                    sb.Append($"[Bảng {tableNum}");
                    sb.Append($" - {table.Rows.Count} hàng x {table.Columns.Count} cột]");

                    // First row as header
                    if (table.Rows.Count > 0)
                    {
                        var headers = new List<string>();
                        foreach (Word.Cell cell in table.Rows[1].Cells)
                        {
                            var text = cell.Range.Text?.Replace("\r\a", "").Trim();
                            if (!string.IsNullOrEmpty(text)) headers.Add(text);
                        }
                        if (headers.Count > 0)
                            sb.Append(" Tiêu đề: " + string.Join(", ", headers));
                    }
                    result.Add(sb.ToString());
                }
            }
            catch { }
            return result;
        }

        public Word.Range GetSelectionRange()
        {
            try
            {
                var sel = _app.Selection;
                if (sel?.Type == Word.WdSelectionType.wdSelectionNormal)
                    return sel.Range;
                return null;
            }
            catch { return null; }
        }

        private Word.Document GetActiveDoc()
        {
            try { return _app.ActiveDocument; }
            catch { return null; }
        }
    }
}
