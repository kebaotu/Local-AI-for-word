using Newtonsoft.Json;

namespace LocalDocAI.Models
{
    public class ChatMessage
    {
        [JsonProperty("role")]
        public string Role { get; set; }

        [JsonProperty("content")]
        public string Content { get; set; }

        public static ChatMessage System(string content) => new ChatMessage { Role = "system", Content = content };
        public static ChatMessage User(string content) => new ChatMessage { Role = "user", Content = content };
        public static ChatMessage Assistant(string content) => new ChatMessage { Role = "assistant", Content = content };
    }
}
