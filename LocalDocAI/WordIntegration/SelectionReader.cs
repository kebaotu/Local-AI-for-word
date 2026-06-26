using System;
using Word = Microsoft.Office.Interop.Word;

namespace LocalDocAI.WordIntegration
{
    public class SelectionReader
    {
        public static string GetText(Word.Application app)
        {
            try
            {
                var sel = app.Selection;
                if (sel?.Type == Word.WdSelectionType.wdSelectionNormal)
                    return sel.Range.Text ?? "";
            }
            catch { }
            return "";
        }

        public static Word.Range GetRange(Word.Application app)
        {
            try
            {
                var sel = app.Selection;
                if (sel?.Type == Word.WdSelectionType.wdSelectionNormal)
                    return sel.Range;
            }
            catch { }
            return null;
        }

        public static bool HasSelection(Word.Application app)
        {
            try
            {
                return app.Selection?.Type == Word.WdSelectionType.wdSelectionNormal
                    && (app.Selection.Range.Text?.Length ?? 0) > 0;
            }
            catch { return false; }
        }
    }
}
