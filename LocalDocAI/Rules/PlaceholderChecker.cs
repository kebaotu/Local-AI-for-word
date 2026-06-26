using LocalDocAI.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LocalDocAI.Rules
{
    public class PlaceholderChecker
    {
        private static readonly Regex PlaceholderPattern =
            new Regex(@"\{\{[^}]+\}\}|\[[^\]]{2,50}\]|<[^>]{2,50}>|_+\s*_+");

        public List<Suggestion> Check(string text)
        {
            var results = new List<Suggestion>();
            int id = 0;

            foreach (Match m in PlaceholderPattern.Matches(text))
            {
                // Skip common abbreviations like [1], [i]
                if (Regex.IsMatch(m.Value, @"^\[\d{1,3}\]$")) continue;
                if (Regex.IsMatch(m.Value, @"^\[(?:i|ii|iii|iv|v)\]$", RegexOptions.IgnoreCase)) continue;

                results.Add(new Suggestion
                {
                    Id = $"PLC-{++id:D4}",
                    Type = "placeholder",
                    Severity = SuggestionSeverity.High,
                    RangeText = m.Value,
                    Explanation = $"Placeholder chưa được điền: {m.Value}",
                    Confidence = 0.9,
                    Action = SuggestionAction.Highlight,
                    RequiresUserReview = true,
                    SafeToAutoApply = false,
                    Source = "rule"
                });
            }

            return results;
        }
    }
}
