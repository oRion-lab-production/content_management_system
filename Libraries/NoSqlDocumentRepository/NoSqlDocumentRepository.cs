using Libraries.LibModels;
using Libraries.NoSqlDocumentRepositorys;
using Libraries.Storage;
using Microsoft.Extensions.Logging;
using System.Text;
using System.Text.Json;

namespace Libraries.NoSqlDocumentRepository
{
    public class FileNoSqlDocumentRepository : INoSqlDocumentRepository
    {
        private readonly string _storageDirectory;
        private readonly ILogger<FileNoSqlDocumentRepository> _logger;

        public FileNoSqlDocumentRepository(IStorageLocator storageLocator, ILogger<FileNoSqlDocumentRepository> logger)
        {
            _logger = logger;

            // Documents are written to <content root>/storage/nosql.
            _storageDirectory = storageLocator.NoSqlPath;
            Directory.CreateDirectory(_storageDirectory);
        }

        public async Task<DocumentSaveResult> SaveDocumentsAsync(Guid jobId, IEnumerable<ContentItem> items, string sourceChannel, CancellationToken cancellationToken)
        {
            _logger.LogInformation("[NoSQL STORE] Saving partition documents to: {Path}", _storageDirectory);

            var result = new DocumentSaveResult();
            int index = 0;

            foreach (var item in items)
            {
                cancellationToken.ThrowIfCancellationRequested();
                index++;

                var externalId = item.ExternalId;
                bool hasExternalId = !string.IsNullOrWhiteSpace(externalId);

                // Partition Key: Groups items by Job/Batch ID
                string partitionKey = $"JOB#{jobId}";

                // Slot Key: deterministic per source + external id so re-imports upsert instead of duplicating.
                string slotKey = hasExternalId ? $"EXTID#{externalId}" : $"ROW#{index}";
                string sourceKey = Sanitize(sourceChannel);

                string documentId = hasExternalId
                    ? $"DOC_{sourceKey}_{Sanitize(externalId!)}"
                    : $"JOB_{jobId:N}_ROW_{index}";

                var fileName = $"{documentId}.json";
                var filePath = Path.Combine(_storageDirectory, fileName);

                // Idempotency: if this content (source + external id) was already stored, skip it.
                if (hasExternalId && File.Exists(filePath))
                {
                    result.SkippedCount++;
                    _logger.LogInformation("[NoSQL STORE] Idempotent skip (already stored) -> {Document}", documentId);
                    continue;
                }

                // Metadata envelope + every dynamic field from the source row (any number of columns).
                var nosqlDocument = new Dictionary<string, object>
                {
                    ["PartitionKey"] = partitionKey,
                    ["SlotKey"] = slotKey,
                    ["JobId"] = jobId,
                    ["LineIndex"] = index,
                    ["SourceChannel"] = sourceChannel,
                    ["StoredAtUtc"] = DateTime.UtcNow,
                    ["DocumentVersion"] = 1,
                    ["FieldCount"] = item.Fields.Count,
                    ["Data"] = item.Fields
                };

                var jsonContent = JsonSerializer.Serialize(nosqlDocument, new JsonSerializerOptions { WriteIndented = true });

                await File.WriteAllTextAsync(filePath, jsonContent, cancellationToken);
                result.ImportedCount++;
                _logger.LogInformation("[NoSQL STORE] Upserted document -> Partition: {Partition}, Slot: {Slot}", partitionKey, slotKey);
            }

            return result;
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "UNKNOWN";

            var builder = new StringBuilder(value.Length);
            foreach (var ch in value)
                builder.Append(char.IsLetterOrDigit(ch) ? ch : '_');
            return builder.ToString();
        }
    }
}
