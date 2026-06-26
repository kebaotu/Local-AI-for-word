using System.Collections.Generic;

namespace LocalDocAI.Models
{
    public class ReviewResult
    {
        public List<Suggestion> Suggestions { get; set; } = new List<Suggestion>();
        public string Summary { get; set; }
        public bool Success { get; set; } = true;
        public string ErrorMessage { get; set; }
    }
}
