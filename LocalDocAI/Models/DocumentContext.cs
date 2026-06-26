using System.Collections.Generic;

namespace LocalDocAI.Models
{
    public class DocumentContext
    {
        public string SelectedText { get; set; }
        public string FullText { get; set; }
        public string CurrentSectionText { get; set; }
        public List<string> Headings { get; set; } = new List<string>();
        public List<CommentInfo> Comments { get; set; } = new List<CommentInfo>();
        public List<RevisionInfo> Revisions { get; set; } = new List<RevisionInfo>();
        public List<string> TableSummaries { get; set; } = new List<string>();
        public bool HasSelection => !string.IsNullOrWhiteSpace(SelectedText);
        public string DocumentFileName { get; set; }
    }
}
