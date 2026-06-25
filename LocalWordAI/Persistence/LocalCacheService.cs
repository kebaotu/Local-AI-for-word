using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace LocalWordAI.Persistence
{
    public class LocalCacheService
    {
        private readonly Dictionary<string, (string value, DateTime expires)> _cache
            = new Dictionary<string, (string, DateTime)>();

        private readonly TimeSpan _ttl;

        public LocalCacheService(TimeSpan? ttl = null)
        {
            _ttl = ttl ?? TimeSpan.FromMinutes(10);
        }

        public void Set(string key, string value)
        {
            _cache[key] = (value, DateTime.UtcNow + _ttl);
        }

        public bool TryGet(string key, out string value)
        {
            if (_cache.TryGetValue(key, out var entry) && entry.expires > DateTime.UtcNow)
            {
                value = entry.value;
                return true;
            }
            _cache.Remove(key);
            value = null;
            return false;
        }

        public void Invalidate(string key) => _cache.Remove(key);

        public void Clear() => _cache.Clear();

        public static string HashText(string text)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(text ?? ""));
                return Convert.ToBase64String(bytes).Substring(0, 16);
            }
        }
    }
}
