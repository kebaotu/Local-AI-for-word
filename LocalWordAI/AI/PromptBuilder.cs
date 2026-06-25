using LocalWordAI.Models;
using System.Collections.Generic;
using System.Text;

namespace LocalWordAI.AI
{
    public class PromptBuilder
    {
        public static readonly string SystemPrompt = @"Bạn là trợ lý AI hữu ích, chạy cục bộ, hỗ trợ người dùng trong mọi tác vụ.
Bạn đặc biệt giỏi về soạn thảo, sửa chữa, rà soát và redline tài liệu trong Microsoft Word.
Trả lời bằng tiếng Việt trừ khi người dùng dùng ngôn ngữ khác.

Khi làm việc với tài liệu:
- Không bịa dữ kiện.
- Không thay đổi ý nghĩa nếu không được yêu cầu.
- Khi đề xuất sửa tài liệu, trả về JSON hợp lệ theo schema yêu cầu.
- Các lỗi pháp lý, logic, số liệu chỉ cảnh báo khi chắc chắn.
- Giữ nguyên formatting, numbering, style nếu có thể.";

        public static readonly string ChatWithActionsSystemPrompt = @"Bạn là trợ lý AI hữu ích chạy cục bộ trong Microsoft Word. Bạn có thể thực hiện các thao tác trực tiếp trên tài liệu Word đang mở.

LUÔN trả về JSON theo format sau (không có markdown code block):
{
  ""message"": ""câu trả lời/giải thích ngắn gọn cho người dùng"",
  ""actions"": []
}

Khi cần thực hiện thao tác trên tài liệu, thêm vào mảng ""actions"":

1. Chèn text tại vị trí con trỏ hiện tại:
{""type"": ""insert_at_cursor"", ""text"": ""nội dung cần chèn"", ""description"": ""mô tả ngắn""}

2. Thêm text vào cuối tài liệu:
{""type"": ""insert_at_end"", ""text"": ""nội dung cần thêm"", ""description"": ""mô tả ngắn""}

3. Tìm đoạn văn và thay thế có track change (để lại dấu vết chỉnh sửa):
{""type"": ""track_change_replace"", ""find"": ""đoạn văn gốc cần tìm (trích nguyên văn)"", ""replacement"": ""nội dung thay thế"", ""description"": ""mô tả ngắn""}

4. Tìm đoạn văn và thay thế trực tiếp (không track change):
{""type"": ""direct_replace"", ""find"": ""đoạn văn gốc"", ""replacement"": ""nội dung thay thế"", ""description"": ""mô tả ngắn""}

5. Thêm comment vào đoạn văn:
{""type"": ""add_comment"", ""find"": ""đoạn văn cần comment"", ""comment"": ""nội dung comment"", ""description"": ""mô tả ngắn""}

6. Highlight đoạn văn:
{""type"": ""highlight"", ""find"": ""đoạn văn cần highlight"", ""description"": ""mô tả ngắn""}

Ví dụ - người dùng yêu cầu dịch vùng chọn sang tiếng Pháp với track change:
{
  ""message"": ""Đã dịch đoạn văn sang tiếng Pháp và áp dụng track change."",
  ""actions"": [{""type"": ""track_change_replace"", ""find"": ""[đoạn văn gốc]"", ""replacement"": ""[bản dịch tiếng Pháp]"", ""description"": ""Dịch sang tiếng Pháp""}]
}

Ví dụ - người dùng yêu cầu viết truyện và thêm vào file:
{
  ""message"": ""Đã viết truyện cười và thêm vào cuối tài liệu."",
  ""actions"": [{""type"": ""insert_at_end"", ""text"": ""[nội dung truyện cười]"", ""description"": ""Thêm truyện cười""}]
}

Nếu chỉ trả lời câu hỏi thông thường, để actions là mảng rỗng [].
Luôn trả lời bằng tiếng Việt trừ khi người dùng dùng ngôn ngữ khác.";

        private static readonly string SuggestionSchema = @"{
  ""suggestions"": [
    {
      ""type"": ""spelling|punctuation|spacing|repeated_word|terminology|capitalization|date_logic|numbering|cross_reference|numeric|passive_voice|legal_risk|placeholder"",
      ""severity"": ""low|medium|high|critical"",
      ""rangeText"": ""đoạn văn cần sửa (trích nguyên văn)"",
      ""replacementText"": ""văn bản thay thế (nếu có)"",
      ""explanation"": ""giải thích ngắn gọn"",
      ""confidence"": 0.0,
      ""action"": ""replace|insert|delete|comment|highlight|none"",
      ""requiresUserReview"": true,
      ""safeToAutoApply"": false
    }
  ]
}";

        public List<ChatMessage> BuildReviewPrompt(string text, string instruction = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Hãy kiểm tra văn bản sau và trả về JSON.");
            sb.AppendLine();
            sb.AppendLine("Cần phát hiện:");
            sb.AppendLine("- chính tả, dấu câu, khoảng trắng, lặp từ");
            sb.AppendLine("- thuật ngữ không thống nhất, tên riêng không thống nhất");
            sb.AppendLine("- viết hoa/viết thường không thống nhất");
            sb.AppendLine("- logic ngày tháng, thực thể mâu thuẫn");
            sb.AppendLine("- đánh số sai, tham chiếu lỗi");
            sb.AppendLine("- số liệu không nhất quán");
            sb.AppendLine("- câu bị động/mơ hồ trách nhiệm");
            sb.AppendLine("- rủi ro pháp lý");
            sb.AppendLine("- placeholder chưa điền");
            if (!string.IsNullOrEmpty(instruction))
            {
                sb.AppendLine();
                sb.AppendLine("Yêu cầu bổ sung: " + instruction);
            }
            sb.AppendLine();
            sb.AppendLine("Schema JSON trả về:");
            sb.AppendLine(SuggestionSchema);
            sb.AppendLine();
            sb.AppendLine("Văn bản:");
            sb.AppendLine("\"\"\"");
            sb.AppendLine(text);
            sb.AppendLine("\"\"\"");

            return new List<ChatMessage>
            {
                ChatMessage.System(SystemPrompt),
                ChatMessage.User(sb.ToString())
            };
        }

        public List<ChatMessage> BuildRewritePrompt(string text, string instruction)
        {
            var userMsg = $@"Hãy viết lại đoạn văn sau theo yêu cầu: {instruction}

Quy tắc:
- Giữ nguyên ý chính.
- Không thêm dữ kiện.
- Không làm thay đổi trách nhiệm pháp lý nếu là văn bản hợp đồng.

Trả về JSON:
{{
  ""originalText"": ""..."",
  ""rewrittenText"": ""..."",
  ""explanation"": ""...""
}}

Văn bản gốc:
---
{text}
---";

            return new List<ChatMessage>
            {
                ChatMessage.System(SystemPrompt),
                ChatMessage.User(userMsg)
            };
        }

        public List<ChatMessage> BuildCommentPrompt(string commentsJson)
        {
            var userMsg = $@"Bạn đang xử lý comments trong tài liệu Word.

Với mỗi comment:
- Tóm tắt yêu cầu.
- Đề xuất phản hồi.
- Đề xuất sửa văn bản nếu cần.
- Phân loại mức độ ưu tiên: high/medium/low.

Trả về JSON:
{{
  ""comments"": [
    {{
      ""id"": ""..."",
      ""category"": ""content_change|reply_needed|verify_needed|user_decision"",
      ""priority"": ""high|medium|low"",
      ""suggestedReply"": ""..."",
      ""suggestedEdit"": ""...""
    }}
  ]
}}

Danh sách comments:
{commentsJson}";

            return new List<ChatMessage>
            {
                ChatMessage.System(SystemPrompt),
                ChatMessage.User(userMsg)
            };
        }

        public List<ChatMessage> BuildRevisionPrompt(string revisionsJson)
        {
            var userMsg = $@"Bạn đang phân tích tracked changes trong tài liệu Word.

Hãy:
- Tóm tắt thay đổi chính.
- Phân loại từng thay đổi: formatting_only|minor|substantive|risky|dealbreaker.
- Phát hiện rủi ro cao.
- Phát hiện dealbreaker nếu có.
- Đề xuất phản hồi/redline ngược lại.

Trả về JSON:
{{
  ""summary"": ""..."",
  ""overallRisk"": ""low|medium|high"",
  ""revisions"": [
    {{
      ""id"": ""..."",
      ""category"": ""formatting_only|minor|substantive|risky|dealbreaker"",
      ""riskLevel"": ""low|medium|high"",
      ""summary"": ""..."",
      ""suggestedResponse"": ""...""
    }}
  ]
}}

Tracked changes:
{revisionsJson}";

            return new List<ChatMessage>
            {
                ChatMessage.System(SystemPrompt),
                ChatMessage.User(userMsg)
            };
        }

        public List<ChatMessage> BuildChatWithActionsMessages(
            string userMessage, DocumentContext context, List<ChatMessage> history = null)
        {
            var messages = new List<ChatMessage> { ChatMessage.System(ChatWithActionsSystemPrompt) };

            // Add history (up to 8 turns)
            if (history != null && history.Count > 0)
            {
                var slice = history.Count > 8 ? history.GetRange(history.Count - 8, 8) : history;
                messages.AddRange(slice);
            }

            // Build user message with context
            var sb = new StringBuilder();
            if (context.HasSelection)
            {
                sb.AppendLine("[Vùng đang chọn trong Word]");
                sb.AppendLine(context.SelectedText);
                sb.AppendLine();
            }
            else if (!string.IsNullOrEmpty(context.FullText))
            {
                var preview = context.FullText.Length > 2000
                    ? context.FullText.Substring(0, 2000) + "\n...(còn tiếp)..."
                    : context.FullText;
                sb.AppendLine("[Nội dung tài liệu hiện tại]");
                sb.AppendLine(preview);
                sb.AppendLine();
            }
            sb.AppendLine("[Yêu cầu của người dùng]");
            sb.AppendLine(userMessage);

            messages.Add(ChatMessage.User(sb.ToString()));
            return messages;
        }

        public List<ChatMessage> BuildChatPrompt(string userMessage, DocumentContext context)
        {
            var sb = new StringBuilder();
            if (context.HasSelection)
            {
                sb.AppendLine("[Vùng chọn hiện tại]");
                sb.AppendLine(context.SelectedText);
                sb.AppendLine();
            }
            else if (!string.IsNullOrEmpty(context.FullText))
            {
                var preview = context.FullText.Length > 3000
                    ? context.FullText.Substring(0, 3000) + "\n...(còn tiếp)..."
                    : context.FullText;
                sb.AppendLine("[Nội dung tài liệu]");
                sb.AppendLine(preview);
                sb.AppendLine();
            }
            sb.AppendLine("[Câu hỏi]");
            sb.AppendLine(userMessage);

            return new List<ChatMessage>
            {
                ChatMessage.System(SystemPrompt),
                ChatMessage.User(sb.ToString())
            };
        }

        public List<ChatMessage> BuildSkillPrompt(string skillPrompt, DocumentContext context)
        {
            var sb = new StringBuilder();
            if (context.HasSelection)
            {
                sb.AppendLine("[Vùng chọn]");
                sb.AppendLine(context.SelectedText);
                sb.AppendLine();
            }
            else if (!string.IsNullOrEmpty(context.FullText))
            {
                sb.AppendLine("[Tài liệu]");
                sb.AppendLine(context.FullText.Length > 4000
                    ? context.FullText.Substring(0, 4000) + "..."
                    : context.FullText);
                sb.AppendLine();
            }

            if (context.Comments.Count > 0)
            {
                sb.AppendLine("[Comments]");
                foreach (var c in context.Comments)
                    sb.AppendLine($"- {c.Author}: {c.Text} (về: {c.ScopeText})");
                sb.AppendLine();
            }

            sb.AppendLine("[Yêu cầu]");
            sb.AppendLine(skillPrompt);

            return new List<ChatMessage>
            {
                ChatMessage.System(SystemPrompt),
                ChatMessage.User(sb.ToString())
            };
        }
    }
}
