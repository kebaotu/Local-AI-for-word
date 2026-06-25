using LocalWordAI.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LocalWordAI.Rules
{
    public class RepeatedWordChecker
    {
        public List<Suggestion> Check(string text)
        {
            var results = new List<Suggestion>();
            int id = 0;

            // Consecutive repeated words (case-insensitive)
            foreach (Match m in Regex.Matches(text, @"\b(\w{2,})\s+\1\b", RegexOptions.IgnoreCase))
            {
                results.Add(new Suggestion
                {
                    Id = $"RPT-{++id:D4}",
                    Type = "repeated_word",
                    Severity = SuggestionSeverity.Medium,
                    RangeText = m.Value,
                    ReplacementText = m.Groups[1].Value,
                    Explanation = $"Từ '{m.Groups[1].Value}' lặp liên tiếp.",
                    Confidence = 0.95,
                    Action = SuggestionAction.Replace,
                    SafeToAutoApply = false,
                    RequiresUserReview = true,
                    Source = "rule"
                });
            }

            return results;
        }
    }
}
