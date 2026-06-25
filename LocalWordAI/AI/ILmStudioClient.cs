using LocalWordAI.Models;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace LocalWordAI.AI
{
    public interface ILmStudioClient
    {
        Task<bool> TestConnectionAsync(CancellationToken ct = default);
        Task<string> ChatAsync(List<ChatMessage> messages, CancellationToken ct = default);
        Task<T> ChatJsonAsync<T>(List<ChatMessage> messages, CancellationToken ct = default) where T : class, new();
        Task<List<string>> ListModelsAsync(CancellationToken ct = default);
    }
}
