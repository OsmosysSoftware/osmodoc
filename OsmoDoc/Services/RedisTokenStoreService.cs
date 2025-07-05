using System;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using StackExchange.Redis;

namespace OsmoDoc.Services;

public class RedisTokenStoreService : IRedisTokenStoreService
{
    private readonly IDatabase _db;
    private const string KeyPrefix = "valid_token:";

    public RedisTokenStoreService(IConnectionMultiplexer redis)
    {
        this._db = redis.GetDatabase();
    }

    public Task StoreTokenAsync(string token, string email, CancellationToken cancellationToken = default)
    {
        // Check if operation was cancelled before starting
        cancellationToken.ThrowIfCancellationRequested();

        return this._db.StringSetAsync($"{KeyPrefix}{token}", JsonConvert.SerializeObject(new
        {
            issuedTo = email,
            issuedAt = DateTime.UtcNow
        }));
    }

    public Task<bool> IsTokenValidAsync(string token, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return this._db.KeyExistsAsync($"{KeyPrefix}{token}");
    }

    public Task RevokeTokenAsync(string token, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return this._db.KeyDeleteAsync($"{KeyPrefix}{token}");
    }
}