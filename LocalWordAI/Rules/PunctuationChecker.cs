using LocalWordAI.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LocalWordAI.Rules
{
    public class PunctuationChecker
    {
        public List<Suggestion> Check(string text)
        {
            var results = new List<Suggestion>();
            int id = 0;

            // Sentence ending without period (lines ending with letter/number but not punctuation)
            // This is heuristic - just flag obvious double periods
            foreach (Match m in Regex.Matches(text, @"\.\s*\."))
            {
                if (m.Value.Trim() != "...") // not ellipsis
                {
                    results.Add(new Suggestion
                    {
                        Id = $"PNC-{++id:D4}",
                        Type = "punctuation",
                        Severity = SuggestionSeverity.Low,
                        RangeText = m.Value,
                        ReplacementText = ".",
                        Explanation = "Dấu chấm kép.",
                        Confidence = 0.9,
                        Action = SuggestionAction.Replace,
                        SafeToAutoApply = false,
                        RequiresUserReview = true,
                        Source = "rule"
                    });
                }
            }

            // Opening quote without closing (simple check)
            var openQuotes = Regex.Matches(text, "\u201C").Count;
            var closeQuotes = Regex.Matches(text, "\u201D").Count;
            if (openQuotes != closeQuotes)
            {
                results.Add(new Suggestion
                {
                    Id = $"PNC-{++id:D4}",
                    Type = "punctuation",
                    Severity = SuggestionSeverity.Medium,
                    RangeText = "",
                    Explanation = $"Số dấu nháy mở ({openQuotes}) không khớp dấu nháy đóng ({closeQuotes}).",
                    Confidence = 0.8,
                    Action = SuggestionAction.Comment,
                    RequiresUserReview = true,
                    SafeToAutoApply = false,
                    Source = "rule"
                });
            }

            return results;
        }
    }
}
