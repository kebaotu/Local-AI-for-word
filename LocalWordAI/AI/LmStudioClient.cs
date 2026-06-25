using LocalWordAI.Models;
using LocalWordAI.Persistence;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Polly;
using Polly.Retry;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace LocalWordAI.AI
{
    public class LmStudioClient : ILmStudioClient
    {
        private readonly HttpClient _http;
        private readonly SettingsService _settings;
        private readonly AsyncRetryPolicy _retryPolicy;

        public LmStudioClient(SettingsService settings)
        {
            _settings = settings;
            _http = new HttpClient();
            _http.Timeout = TimeSpan.FromSeconds(settings.Current.TimeoutSeconds);

            _retryPolicy = Policy
                .Handle<HttpRequestException>()
                .Or<TaskCanceledException>()
                .WaitAndRetryAsync(2, attempt => TimeSpan.FromSeconds(attempt * 2));
        }

        public async Task<bool> TestConnectionAsync(CancellationToken ct = default)
        {
            try
            {
                var url = _settings.Current.LmStudioBaseUrl.TrimEnd('/') + "/v1/models";
                var resp = await _http.GetAsync(url, ct);
                return resp.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        public async Task<List<string>> ListModelsAsync(CancellationToken ct = default)
        {
            var result = new List<string>();
            try
            {
                var url = _settings.Current.LmStudioBaseUrl.TrimEnd('/') + "/v1/models";
                var resp = await _http.GetStringAsync(url);
                var obj = JObject.Parse(resp);
                foreach (var item in obj["data"] ?? new JArray())
                    result.Add(item["id"]?.ToString() ?? "");
            }
            catch { }
            return result;
        }

        public async Task<string> ChatAsync(List<ChatMessage> messages, CancellationToken ct = default)
        {
            return await _retryPolicy.ExecuteAsync(async () =>
            {
                var url = _settings.Current.LmStudioBaseUrl.TrimEnd('/') + "/v1/chat/completions";

                var payload = new
                {
                    model = string.IsNullOrEmpty(_settings.Current.ModelName) ? "local-model" : _settings.Current.ModelName,
                    messages = messages,
                    temperature = _settings.Current.Temperature,
                    max_tokens = _settings.Current.MaxTokens,
                    stream = false
                };

                var json = JsonConvert.SerializeObject(payload);
                using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
                using (var cts = CancellationTokenSource.CreateLinkedTokenSource(ct))
                {
                    cts.CancelAfter(TimeSpan.FromSeconds(_settings.Current.TimeoutSeconds));
                    var response = await _http.PostAsync(url, content, cts.Token);
                    response.EnsureSuccessStatusCode();
                    var responseJson = await response.Content.ReadAsStringAsync();
                    var obj = JObject.Parse(responseJson);
                    return obj["choices"]?[0]?["message"]?["content"]?.ToString() ?? "";
                }
            });
        }

        public async Task<T> ChatJsonAsync<T>(List<ChatMessage> messages, CancellationToken ct = default) where T : class, new()
        {
            var raw = await ChatAsync(messages, ct);
            return JsonOutputParser.Parse<T>(raw);
        }
    }
}
