using LocalDocAI.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LocalDocAI.Rules
{
    public class SpacingChecker
    {
        public List<Suggestion> Check(string text)
        {
            var results = new List<Suggestion>();
            int id = 0;

            // Double (or more) spaces
            foreach (Match m in Regex.Matches(text, @"[ ]{2,}"))
            {
                results.Add(new Suggestion
                {
                    Id = $"SPC-{++id:D4}",
                    Type = "spacing",
                    Severity = SuggestionSeverity.Low,
                    RangeText = m.Value,
                    ReplacementText = " ",
                    Explanation = "Khoảng trắng thừa.",
                    Confidence = 1.0,
                    Action = SuggestionAction.Replace,
                    SafeToAutoApply = true,
                    RequiresUserReview = false,
                    Source = "rule"
                });
            }

            // Space before punctuation
            foreach (Match m in Regex.Matches(text, @" +([,\.;:!?\)])"))
            {
                results.Add(new Suggestion
                {
                    Id = $"SPC-{++id:D4}",
                    Type = "spacing",
                    Severity = SuggestionSeverity.Low,
                    RangeText = m.Value,
                    ReplacementText = m.Groups[1].Value,
                    Explanation = $"Dấu cách thừa trước '{m.Groups[1].Value}'.",
                    Confidence = 0.95,
                    Action = SuggestionAction.Replace,
                    SafeToAutoApply = true,
                    RequiresUserReview = false,
                    Source = "rule"
                });
            }

            // Missing space after punctuation (not decimal numbers)
            foreach (Match m in Regex.Matches(text, @"([,;:])([^\s\d\r\n""'])"))
            {
                results.Add(new Suggestion
                {
                    Id = $"SPC-{++id:D4}",
                    Type = "spacing",
                    Severity = SuggestionSeverity.Low,
                    RangeText = m.Value,
                    ReplacementText = m.Groups[1].Value + " " + m.Groups[2].Value,
                    Explanation = $"Thiếu dấu cách sau '{m.Groups[1].Value}'.",
                    Confidence = 0.85,
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
