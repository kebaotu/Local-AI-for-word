using System;
using Word = Microsoft.Office.Interop.Word;

namespace LocalDocAI.WordIntegration
{
    public class HighlightService
    {
        private readonly Word.Application _app;

        public HighlightService(Word.Application app)
        {
            _app = app;
        }

        public bool HighlightText(Word.Document doc, string searchText,
            Word.WdColorIndex color = Word.WdColorIndex.wdYellow)
        {
            try
            {
                var find = doc.Content.Duplicate;
                find.Find.ClearFormatting();
                find.Find.Text = searchText;
                find.Find.Forward = true;
                find.Find.Wrap = Word.WdFindWrap.wdFindStop;
                if (find.Find.Execute())
                {
                    find.HighlightColorIndex = color;
                    return true;
                }
                return false;
            }
            catch { return false; }
        }

        public bool HighlightRange(Word.Range range, Word.WdColorIndex color = Word.WdColorIndex.wdYellow)
        {
            try
            {
                range.HighlightColorIndex = color;
                return true;
            }
            catch { return false; }
        }

        public Word.WdColorIndex GetColorBySeverity(string severity)
        {
            switch (severity?.ToLower())
            {
                case "critical": return Word.WdColorIndex.wdRed;
                case "high": return Word.WdColorIndex.wdRed;
                case "medium": return Word.WdColorIndex.wdYellow;
                case "low": return Word.WdColorIndex.wdBrightGreen;
                default: return Word.WdColorIndex.wdYellow;
            }
        }
    }
}
