using LocalWordAI.Models;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LocalWordAI.Rules
{
    public class NumberingChecker
    {
        public List<Suggestion> Check(string text)
        {
            var results = new List<Suggestion>();
            int id = 0;

            // Check simple numbered lists: 1. 2. 4. (skip)
            var listItems = new List<(int num, string raw, int pos)>();
            foreach (Match m in Regex.Matches(text, @"^(\d+)\.\s", RegexOptions.Multiline))
            {
                if (int.TryParse(m.Groups[1].Value, out int n))
                    listItems.Add((n, m.Value.Trim(), m.Index));
            }

            for (int i = 1; i < listItems.Count; i++)
            {
                var prev = listItems[i - 1];
                var curr = listItems[i];
                if (curr.num != prev.num + 1 && curr.num > prev.num)
                {
                    results.Add(new Suggestion
                    {
                        Id = $"NUM-{++id:D4}",
                        Type = "numbering",
                        Severity = SuggestionSeverity.Medium,
                        RangeText = curr.raw,
                        Explanation = $"Đánh số bị nhảy từ {prev.num} sang {curr.num}, có thể thiếu mục {prev.num + 1}.",
                        Confidence = 0.85,
                        Action = SuggestionAction.Comment,
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
