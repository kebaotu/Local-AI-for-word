using LocalWordAI.AI;
using LocalWordAI.Models;
using System.Threading;
using System.Threading.Tasks;

namespace LocalWordAI.Skills
{
    public class SkillRunner
    {
        private readonly ILmStudioClient _client;
        private readonly PromptBuilder _promptBuilder;

        public SkillRunner(ILmStudioClient client, PromptBuilder promptBuilder)
        {
            _client = client;
            _promptBuilder = promptBuilder;
        }

        public async Task<ReviewResult> RunAsync(Skill skill, DocumentContext context, CancellationToken ct = default)
        {
            var result = new ReviewResult();
            try
            {
                var promptOverride = string.IsNullOrEmpty(skill.SystemPromptOverride)
                    ? null : skill.SystemPromptOverride;

                var messages = _promptBuilder.BuildSkillPrompt(skill.Prompt, context);
                if (promptOverride != null)
                    messages[0] = ChatMessage.System(promptOverride);

                if (skill.OutputMode == "suggestions")
                {
                    var parsed = await _client.ChatJsonAsync<SuggestionList>(messages, ct);
                    result.Suggestions = parsed?.Suggestions ?? new System.Collections.Generic.List<Suggestion>();
                }
                else
                {
                    result.Summary = await _client.ChatAsync(messages, ct);
                }

                result.Success = true;
            }
            catch (System.Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            return result;
        }
    }
}
