using LocalDocAI.Models;
using System.Collections.Generic;
using System.Text;

namespace LocalDocAI.AI
{
    public class ContextCompressor
    {
        private const int MaxChars = 8000;

        public string CompressForReview(string fullText, string selectedText)
        {
            if (!string.IsNullOrEmpty(selectedText) && selectedText.Length <= MaxChars)
                return selectedText;

            if (fullText.Length <= MaxChars)
                return fullText;

            // Take first chunk if too long
            return fullText.Substring(0, MaxChars) + "\n...[văn bản bị cắt bớt do quá dài]...";
        }

        public List<string> ChunkDocument(string fullText, int chunkSizeChars = 6000)
        {
            var chunks = new List<string>();
            if (string.IsNullOrEmpty(fullText)) return chunks;

            // Split on paragraph breaks
            var paragraphs = fullText.Split(new[] { "\r\n\r\n", "\n\n" }, System.StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();

            foreach (var para in paragraphs)
            {
                if (sb.Length + para.Length > chunkSizeChars && sb.Length > 0)
                {
                    chunks.Add(sb.ToString());
                    sb.Clear();
                }
                sb.AppendLine(para);
            }

            if (sb.Length > 0) chunks.Add(sb.ToString());
            return chunks;
        }
    }
}
