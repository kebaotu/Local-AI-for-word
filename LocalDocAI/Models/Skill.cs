using Newtonsoft.Json;
using System.Collections.Generic;

namespace LocalDocAI.Models
{
    public class Skill
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("context")]
        public List<string> Context { get; set; } = new List<string>();

        [JsonProperty("prompt")]
        public string Prompt { get; set; }

        [JsonProperty("outputMode")]
        public string OutputMode { get; set; } = "suggestions";

        [JsonProperty("defaultAction")]
        public string DefaultAction { get; set; } = "comment";

        [JsonProperty("systemPromptOverride")]
        public string SystemPromptOverride { get; set; }
    }
}
