using LocalWordAI.Models;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LocalWordAI.Rules
{
    public class DateLogicChecker
    {
        private static readonly Regex DatePattern =
            new Regex(@"\b([0-3]?\d)[\/\-\.]([0-1]?\d)[\/\-\.](\d{4})\b");

        public List<Suggestion> Check(string text)
        {
            var results = new List<Suggestion>();
            int id = 0;

            // Find all dates and validate
            var dates = new List<(string raw, DateTime dt, int pos)>();

            foreach (Match m in DatePattern.Matches(text))
            {
                int day, month, year;
                if (!int.TryParse(m.Groups[1].Value, out day)) continue;
                if (!int.TryParse(m.Groups[2].Value, out month)) continue;
                if (!int.TryParse(m.Groups[3].Value, out year)) continue;

                // Invalid date values
                if (day < 1 || day > 31 || month < 1 || month > 12)
                {
                    results.Add(new Suggestion
                    {
                        Id = $"DT-{++id:D4}",
                        Type = "date_logic",
                        Severity = SuggestionSeverity.High,
                        RangeText = m.Value,
                        Explanation = $"Ngày tháng không hợp lệ: {m.Value}",
                        Confidence = 0.95,
                        Action = SuggestionAction.Highlight,
                        RequiresUserReview = true,
                        SafeToAutoApply = false,
                        Source = "rule"
                    });
                    continue;
                }

                try
                {
                    var dt = new DateTime(year, month, day);
                    dates.Add((m.Value, dt, m.Index));
                }
                catch
                {
                    results.Add(new Suggestion
                    {
                        Id = $"DT-{++id:D4}",
                        Type = "date_logic",
                        Severity = SuggestionSeverity.High,
                        RangeText = m.Value,
                        Explanation = $"Ngày không tồn tại: {m.Value} (ví dụ: 30/02 không hợp lệ).",
                        Confidence = 1.0,
                        Action = SuggestionAction.Highlight,
                        RequiresUserReview = true,
                        SafeToAutoApply = false,
                        Source = "rule"
                    });
                }
            }

            // Check for "reply before" date earlier than "issued" date in same sentence
            for (int i = 1; i < dates.Count; i++)
            {
                // Simple heuristic: if date[i] < date[i-1] and they are close in text, flag
                if ((dates[i].pos - dates[i - 1].pos) < 200 && dates[i].dt < dates[i - 1].dt)
                {
                    results.Add(new Suggestion
                    {
                        Id = $"DT-{++id:D4}",
                        Type = "date_logic",
                        Severity = SuggestionSeverity.High,
                        RangeText = $"{dates[i - 1].raw} ... {dates[i].raw}",
                        Explanation = $"Ngày {dates[i].raw} nhỏ hơn ngày {dates[i - 1].raw} đứng trước — có thể lỗi logic.",
                        Confidence = 0.7,
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
