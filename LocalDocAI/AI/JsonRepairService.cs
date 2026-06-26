using System.Text.RegularExpressions;

namespace LocalDocAI.AI
{
    public static class JsonRepairService
    {
        public static string Repair(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return "{}";

            // Remove trailing commas before } or ]
            json = Regex.Replace(json, @",\s*([}\]])", "$1");

            // Close unclosed arrays/objects naively
            int opens = 0, closes = 0, arrOpens = 0, arrCloses = 0;
            foreach (char c in json)
            {
                if (c == '{') opens++;
                else if (c == '}') closes++;
                else if (c == '[') arrOpens++;
                else if (c == ']') arrCloses++;
            }

            while (arrOpens > arrCloses) { json += "]"; arrCloses++; }
            while (opens > closes) { json += "}"; closes++; }

            return json;
        }
    }
}
