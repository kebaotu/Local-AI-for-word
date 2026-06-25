using Newtonsoft.Json;
using System.Collections.Generic;

namespace LocalWordAI.Models
{
    public class ChatActionResponse
    {
        [JsonProperty("message")]
        public string Message { get; set; }

        [JsonProperty("actions")]
        public List<ChatAction> Actions { get; set; } = new List<ChatAction>();
    }

    public class ChatAction
    {
        // insert_at_cursor | insert_at_end | track_change_replace | direct_replace | add_comment | highlight
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        // For insert_at_cursor / insert_at_end
        [JsonProperty("text")]
        public string Text { get; set; }

        // For track_change_replace / direct_replace / add_comment / highlight
        [JsonProperty("find")]
        public string Find { get; set; }

        [JsonProperty("replacement")]
        public string Replacement { get; set; }

        [JsonProperty("comment")]
        public string Comment { get; set; }
    }
}
