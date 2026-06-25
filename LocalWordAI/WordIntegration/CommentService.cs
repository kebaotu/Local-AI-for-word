using LocalWordAI.Persistence;
using System;
using Word = Microsoft.Office.Interop.Word;

namespace LocalWordAI.WordIntegration
{
    public class CommentService
    {
        private readonly Word.Application _app;
        private readonly SettingsService _settings;
        private const string AiPrefix = "[Local AI] ";

        public CommentService(Word.Application app, SettingsService settings)
        {
            _app = app;
            _settings = settings;
        }

        public bool AddComment(Word.Document doc, Word.Range range, string commentText)
        {
            try
            {
                doc.Comments.Add(range, AiPrefix + commentText);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("AddComment error: " + ex.Message);
                return false;
            }
        }

        public bool AddCommentOnText(Word.Document doc, string searchText, string commentText)
        {
            try
            {
                var find = doc.Content.Duplicate;
                find.Find.ClearFormatting();
                find.Find.Text = searchText;
                find.Find.Forward = true;
                find.Find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Find.MatchCase = true;
                if (find.Find.Execute())
                {
                    doc.Comments.Add(find, AiPrefix + commentText);
                    return true;
                }
                return false;
            }
            catch { return false; }
        }

        public int RemoveAiComments(Word.Document doc)
        {
            int count = 0;
            try
            {
                // Collect indices first to avoid modifying collection while iterating
                var toDelete = new System.Collections.Generic.List<Word.Comment>();
                foreach (Word.Comment c in doc.Comments)
                {
                    if ((c.Range?.Text ?? "").StartsWith(AiPrefix))
                        toDelete.Add(c);
                }
                foreach (var c in toDelete)
                {
                    c.Delete();
                    count++;
                }
            }
            catch { }
            return count;
        }
    }
}
