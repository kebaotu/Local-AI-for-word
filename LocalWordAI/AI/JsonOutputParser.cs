using Newtonsoft.Json;
using System;
using System.Text.RegularExpressions;

namespace LocalWordAI.AI
{
    public static class JsonOutputParser
    {
        public static T Parse<T>(string raw) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(raw)) return new T();

            // Try direct parse first
            try
            {
                return JsonConvert.DeserializeObject<T>(raw.Trim()) ?? new T();
            }
            catch { }

            // Extract JSON from markdown code block
            var extracted = ExtractJsonFromMarkdown(raw);
            if (!string.IsNullOrEmpty(extracted))
            {
                try { return JsonConvert.DeserializeObject<T>(extracted) ?? new T(); }
                catch { }
            }

            // Try JSON repair
            var repaired = JsonRepairService.Repair(extracted ?? raw);
            try { return JsonConvert.DeserializeObject<T>(repaired) ?? new T(); }
            catch { return new T(); }
        }

        private static string ExtractJsonFromMarkdown(string text)
        {
            // Match ```json ... ``` or ``` ... ```
            var match = Regex.Match(text, @"```(?:json)?\s*(\{[\s\S]*?\})\s*```", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value.Trim();

            // Match bare { ... } at outermost level
            int start = text.IndexOf('{');
            int end = text.LastIndexOf('}');
            if (start >= 0 && end > start)
                return text.Substring(start, end - start + 1);

            return null;
        }
    }
}
