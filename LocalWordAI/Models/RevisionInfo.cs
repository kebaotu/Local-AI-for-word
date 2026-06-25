using Newtonsoft.Json;

namespace LocalWordAI.Models
{
    public class RevisionInfo
    {
        public string Id { get; set; }
        public string Author { get; set; }
        public string Type { get; set; }
        public string Text { get; set; }
        public string Date { get; set; }
        public string RiskLevel { get; set; }
        public string Summary { get; set; }
        public string Category { get; set; }
        public int RangeStart { get; set; }
        public int RangeEnd { get; set; }
    }
}
