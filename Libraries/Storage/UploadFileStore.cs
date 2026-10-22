using Microsoft.Extensions.Logging;
using System.Text;

namespace Libraries.Storage
{
    public class UploadFileStore : IUploadFileStore
    {
        private readonly IStorageLocator _locator;
        private readonly ILogger<UploadFileStore> _logger;

        public UploadFileStore(IStorageLocator locator, ILogger<UploadFileStore> logger)
        {
            _locator = locator;
            _logger = logger;
        }

        public async Task<string> SaveAsync(Guid jobId, string fileName, byte[] content, CancellationToken cancellationToken)
        {
            var safeName = Sanitize(Path.GetFileName(fileName));
            if (string.IsNullOrWhiteSpace(safeName))
                safeName = "upload.bin";

            // Prefix with the job id so re-uploads never overwrite a previous submission on disk.
            var storedName = $"{jobId:N}_{safeName}";
            var path = Path.Combine(_locator.FilePath, storedName);

            await File.WriteAllBytesAsync(path, content, cancellationToken);
            _logger.LogInformation("[FILE STORE] Saved raw upload -> {Path}", path);

            return path;
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var builder = new StringBuilder(value.Length);
            foreach (var ch in value)
                builder.Append(char.IsLetterOrDigit(ch) || ch == '.' || ch == '-' || ch == '_' ? ch : '_');
            return builder.ToString();
        }
    }
}
