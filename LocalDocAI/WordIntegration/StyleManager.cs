using System;
using Word = Microsoft.Office.Interop.Word;

namespace LocalDocAI.WordIntegration
{
    public class StyleManager
    {
        public static object CaptureStyle(Word.Range range)
        {
            try { return range.get_Style(); }
            catch { return null; }
        }

        public static void RestoreStyle(Word.Range range, object style)
        {
            if (style == null) return;
            try { range.set_Style(style); }
            catch { }
        }

        public static bool IsInTable(Word.Range range)
        {
            try { return range.Tables.Count > 0; }
            catch { return false; }
        }

        public static bool IsInList(Word.Range range)
        {
            try { return range.ListFormat?.ListType != Word.WdListType.wdListNoNumbering; }
            catch { return false; }
        }
    }
}
