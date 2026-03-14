using Microsoft.Extensions.Caching.Distributed;
using System.Text.Json;

namespace MOS.ExcelGrading.API.Helpers
{
    internal static class DistributedCacheJsonExtensions
    {
        public static async Task<T?> GetJsonAsync<T>(this IDistributedCache cache, string key, CancellationToken cancellationToken = default)
        {
            var payload = await cache.GetStringAsync(key, cancellationToken);
            if (string.IsNullOrWhiteSpace(payload))
                return default;

            try
            {
                return JsonSerializer.Deserialize<T>(payload);
            }
            catch
            {
                return default;
            }
        }

        public static Task SetJsonAsync<T>(
            this IDistributedCache cache,
            string key,
            T value,
            TimeSpan ttl,
            CancellationToken cancellationToken = default)
        {
            var payload = JsonSerializer.Serialize(value);
            return cache.SetStringAsync(
                key,
                payload,
                new DistributedCacheEntryOptions
                {
                    AbsoluteExpirationRelativeToNow = ttl
                },
                cancellationToken);
        }
    }
}
