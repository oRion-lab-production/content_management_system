using Libraries.Storage;
using Microsoft.Extensions.Logging;
using System.Security.Cryptography;
using System.Text;

namespace Libraries.Idempotency
{
    public class FileIdempotencyStore : IIdempotencyStore
    {
        private readonly string _directory;
        private readonly ILogger<FileIdempotencyStore> _logger;

        public FileIdempotencyStore(IStorageLocator storageLocator, ILogger<FileIdempotencyStore> logger)
        {
            _logger = logger;

            _directory = Path.Combine(storageLocator.NoSqlPath, "idempotency");
            Directory.CreateDirectory(_directory);
        }

        public async Task<Guid?> TryGetJobIdAsync(string key, CancellationToken cancellationToken)
        {
            var path = GetPath(key);
            if (!File.Exists(path))
                return null;

            var content = await File.ReadAllTextAsync(path, cancellationToken);
            return Guid.TryParse(content, out var jobId) ? jobId : (Guid?)null;
        }

        public async Task SaveAsync(string key, Guid jobId, CancellationToken cancellationToken)
        {
            var path = GetPath(key);
            await File.WriteAllTextAsync(path, jobId.ToString(), cancellationToken);
            _logger.LogInformation("[IDEMPOTENCY] Recorded key {Key} -> Job {JobId}", Shorten(key), jobId);
        }

        private string GetPath(string key)
        {
            using var sha = SHA256.Create();
            var hash = Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(key)));
            return Path.Combine(_directory, $"{hash}.json");
        }

        private static string Shorten(string key) =>
            key.Length <= 12 ? key : key.Substring(0, 12) + "...";
    }
}
