using Newtonsoft.Json;
using System.Collections.Generic;

namespace LocalDocAI.Models
{
    public enum SuggestionSeverity { Low, Medium, High, Critical }
    public enum SuggestionStatus { Pending, Accepted, Rejected, Ignored }
    public enum SuggestionAction { Replace, Insert, Delete, Comment, Highlight, None }

    public class Suggestion
    {
        public string Id { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("severity")]
        public string SeverityStr
        {
            get => Severity.ToString().ToLower();
            set
            {
                if (System.Enum.TryParse<SuggestionSeverity>(value, true, out var s))
                    Severity = s;
            }
        }

        [JsonIgnore]
        public SuggestionSeverity Severity { get; set; } = SuggestionSeverity.Medium;

        [JsonProperty("rangeText")]
        public string RangeText { get; set; }

        [JsonProperty("replacementText")]
        public string ReplacementText { get; set; }

        [JsonProperty("explanation")]
        public string Explanation { get; set; }

        [JsonProperty("confidence")]
        public double Confidence { get; set; } = 0.8;

        [JsonProperty("action")]
        public string ActionStr
        {
            get => Action.ToString().ToLower();
            set
            {
                if (System.Enum.TryParse<SuggestionAction>(value, true, out var a))
                    Action = a;
            }
        }

        [JsonIgnore]
        public SuggestionAction Action { get; set; } = SuggestionAction.Comment;

        [JsonProperty("requiresUserReview")]
        public bool RequiresUserReview { get; set; } = true;

        [JsonProperty("safeToAutoApply")]
        public bool SafeToAutoApply { get; set; } = false;

        [JsonProperty("contextBefore")]
        public string ContextBefore { get; set; }

        [JsonProperty("contextAfter")]
        public string ContextAfter { get; set; }

        [JsonProperty("source")]
        public string Source { get; set; } = "llm";

        [JsonIgnore]
        public SuggestionStatus Status { get; set; } = SuggestionStatus.Pending;

        // Word range start/end for applying
        [JsonIgnore]
        public int RangeStart { get; set; } = -1;
        [JsonIgnore]
        public int RangeEnd { get; set; } = -1;
    }

    public class SuggestionList
    {
        [JsonProperty("suggestions")]
        public List<Suggestion> Suggestions { get; set; } = new List<Suggestion>();
    }
}
