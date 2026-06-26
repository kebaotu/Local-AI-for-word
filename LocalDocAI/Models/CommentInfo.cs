using Newtonsoft.Json;

namespace LocalDocAI.Models
{
    public class CommentInfo
    {
        public string Id { get; set; }
        public string Author { get; set; }
        public string Text { get; set; }
        public string ScopeText { get; set; }
        public int Start { get; set; }
        public int End { get; set; }
        public string Date { get; set; }

        // AI analysis fields
        public string Category { get; set; }
        public string Priority { get; set; }
        public string SuggestedReply { get; set; }
        public string SuggestedEdit { get; set; }
    }
}
