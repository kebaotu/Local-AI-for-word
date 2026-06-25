using LocalWordAI.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LocalWordAI.Rules
{
    public class CrossReferenceChecker
    {
        public List<Suggestion> Check(string text)
        {
            var results = new List<Suggestion>();
            int id = 0;

            // Parse defined sections: "Điều X", "Khoản X", "Mục X", "Phụ lục X", "Bảng X"
            var defined = new HashSet<string>();
            foreach (Match m in Regex.Matches(text,
                @"\b(Điều|Khoản|Mục|Phụ lục|Bảng|Hình)\s+(\d+|[IVXivx]+|[A-Za-z])\b"))
            {
                defined.Add(m.Value.Trim());
            }

            // Find references in text
            foreach (Match m in Regex.Matches(text,
                @"\btheo\s+(Điều|Khoản|Mục|Phụ lục|Bảng|Hình)\s+(\d+|[IVXivx]+|[A-Za-z])\b"))
            {
                var refStr = m.Value.Replace("theo ", "").Trim();
                // Normalize: "Điều 1" vs "điều 1"
                bool found = false;
                foreach (var d in defined)
                {
                    if (string.Compare(d, refStr, true) == 0) { found = true; break; }
                }

                if (!found)
                {
                    results.Add(new Suggestion
                    {
                        Id = $"XRF-{++id:D4}",
                        Type = "cross_reference",
                        Severity = SuggestionSeverity.High,
                        RangeText = m.Value,
                        Explanation = $"Tham chiếu '{refStr}' không tìm thấy trong tài liệu.",
                        Confidence = 0.75,
                        Action = SuggestionAction.Highlight,
                        RequiresUserReview = true,
                        SafeToAutoApply = false,
                        Source = "rule"
                    });
                }
            }

            return results;
        }
    }
}
