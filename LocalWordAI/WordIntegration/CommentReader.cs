using LocalWordAI.Models;
using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace LocalWordAI.WordIntegration
{
    public static class CommentReader
    {
        public static List<CommentInfo> ReadAll(Word.Document doc)
        {
            var result = new List<CommentInfo>();
            try
            {
                int i = 0;
                foreach (Word.Comment c in doc.Comments)
                {
                    i++;
                    result.Add(new CommentInfo
                    {
                        Id = "CMT-" + i.ToString("D4"),
                        Author = c.Author ?? "Unknown",
                        Text = c.Range?.Text?.TrimEnd('\r') ?? "",
                        ScopeText = c.Scope?.Text?.TrimEnd('\r') ?? "",
                        Start = c.Scope?.Start ?? 0,
                        End = c.Scope?.End ?? 0,
                        Date = c.Date.ToString("yyyy-MM-dd HH:mm")
                    });
                }
            }
            catch { }
            return result;
        }
    }
}
