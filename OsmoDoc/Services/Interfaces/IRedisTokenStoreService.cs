using System.Threading;
using System.Threading.Tasks;

namespace OsmoDoc.Services;

public interface IRedisTokenStoreService
{
    Task StoreTokenAsync(string token, string email, CancellationToken cancellationToken = default);
    Task<bool> IsTokenValidAsync(string token, CancellationToken cancellationToken = default);
    Task RevokeTokenAsync(string token, CancellationToken cancellationToken = default);
}