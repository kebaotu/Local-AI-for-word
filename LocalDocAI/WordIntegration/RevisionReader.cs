using LocalDocAI.Models;
using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace LocalDocAI.WordIntegration
{
    public static class RevisionReader
    {
        public static List<RevisionInfo> ReadAll(Word.Document doc)
        {
            var result = new List<RevisionInfo>();
            try
            {
                int i = 0;
                foreach (Word.Revision rev in doc.Revisions)
                {
                    i++;
                    result.Add(new RevisionInfo
                    {
                        Id = "REV-" + i.ToString("D4"),
                        Author = rev.Author ?? "Unknown",
                        Type = GetRevisionType(rev.Type),
                        Text = rev.Range?.Text?.TrimEnd('\r') ?? "",
                        Date = rev.Date.ToString("yyyy-MM-dd HH:mm"),
                        RangeStart = rev.Range?.Start ?? 0,
                        RangeEnd = rev.Range?.End ?? 0
                    });
                }
            }
            catch { }
            return result;
        }

        private static string GetRevisionType(Word.WdRevisionType type)
        {
            switch (type)
            {
                case Word.WdRevisionType.wdRevisionInsert: return "insert";
                case Word.WdRevisionType.wdRevisionDelete: return "delete";
                case Word.WdRevisionType.wdRevisionProperty: return "format";
                case Word.WdRevisionType.wdRevisionStyle: return "style";
                case Word.WdRevisionType.wdRevisionReplace: return "replace";
                default: return "other";
            }
        }
    }
}
