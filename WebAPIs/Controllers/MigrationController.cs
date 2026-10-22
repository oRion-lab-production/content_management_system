using Libraries.BackgroundProcess;
using Libraries.ExcelParser;
using Libraries.Idempotency;
using Libraries.LibModels;
using Libraries.Storage;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using System.Security.Cryptography;
using System.Text.Json;

namespace WebAPIs.Controllers
{
    public class MigrationController : APIBase
    {
        private readonly IBackgroundQueue _queue;
        private readonly IExcelParserService _excelParser;
        private readonly IIdempotencyStore _idempotency;
        private readonly IUploadFileStore _fileStore;
        private readonly MigrationThresholds _thresholds;

        public MigrationController(
            IBackgroundQueue queue,
            IExcelParserService excelParser,
            IIdempotencyStore idempotency,
            IUploadFileStore fileStore,
            IOptions<MigrationThresholds> thresholds)
        {
            _queue = queue;
            _excelParser = excelParser;
            _idempotency = idempotency;
            _fileStore = fileStore;
            _thresholds = thresholds.Value;
        }

        [HttpPost]
        public async Task<IActionResult> PushContent(
            [FromBody] DirectJsonImportRequest request,
            [FromHeader(Name = "Idempotency-Key")] string idempotencyKey,
            CancellationToken cancellationToken)
        {
            if (request.Items == null || !request.Items.Any())
                return BadRequest("Payload items cannot be empty.");

            var payloadBytes = JsonSerializer.SerializeToUtf8Bytes(request);

            // Idempotency: if the same submission (or key) was already accepted, return the original job.
            var payloadKey = string.IsNullOrWhiteSpace(idempotencyKey)
                ? "json:" + ComputeHash(payloadBytes)
                : idempotencyKey;

            var existingJobId = await _idempotency.TryGetJobIdAsync(payloadKey, cancellationToken);
            if (existingJobId.HasValue)
            {
                return Ok(new
                {
                    JobId = existingJobId.Value,
                    Status = "Duplicate",
                    Message = "This payload was already accepted; returning the original job."
                });
            }

            if (request.Items.Count > _thresholds.MaxJsonItems)
            {
                return BadRequest(new ProblemDetails
                {
                    Status = StatusCodes.Status400BadRequest,
                    Title = "Threshold Exceeded",
                    Detail = $"Payload contains {request.Items.Count} items, exceeding the maximum allowed limit of {_thresholds.MaxJsonItems} items."
                });
            }

            // Idempotency within the batch: drop repeated ExternalId values.
            var dedupedItems = DeduplicateByExternalId(request.Items);

            var job = new MigrationJob
            {
                JobId = Guid.NewGuid(),
                SourceChannel = "JSON_API",
                Items = dedupedItems,
                NotificationEmail = null,
                CreatedAtUtc = DateTime.UtcNow
            };

            // Persist the raw JSON payload to storage/file for traceability.
            var sourceName = string.IsNullOrWhiteSpace(request.SourceSystem) ? "content" : request.SourceSystem;
            await _fileStore.SaveAsync(job.JobId, $"{sourceName}.json", payloadBytes, cancellationToken);

            await _queue.EnqueueAsync(job, cancellationToken);
            await _idempotency.SaveAsync(payloadKey, job.JobId, cancellationToken);

            return NoContent();
        }

        [HttpPost]
        public async Task<IActionResult> UploadExcel(
            IFormFile file,
            [FromQuery] string notificationEmail,
            [FromHeader(Name = "Idempotency-Key")] string idempotencyKey,
            CancellationToken cancellationToken)
        {
            if (file == null || file.Length == 0)
                return BadRequest("A valid file upload is required.");

            long maxSizeBytes = (long)_thresholds.MaxExcelSizeMb * 1024 * 1024;
            if (file.Length > maxSizeBytes)
            {
                return StatusCode(StatusCodes.Status413PayloadTooLarge, new ProblemDetails
                {
                    Status = StatusCodes.Status413PayloadTooLarge,
                    Title = "File Size Exceeded",
                    Detail = $"File size exceeds maximum threshold of {_thresholds.MaxExcelSizeMb} MB."
                });
            }

            // Buffer the upload so we can hash it (idempotency) and then parse it as a seekable stream.
            using var buffer = new MemoryStream();
            await file.CopyToAsync(buffer, cancellationToken);
            var fileBytes = buffer.ToArray();

            var payloadKey = string.IsNullOrWhiteSpace(idempotencyKey)
                ? "excel:" + ComputeHash(fileBytes)
                : idempotencyKey;

            var existingJobId = await _idempotency.TryGetJobIdAsync(payloadKey, cancellationToken);
            if (existingJobId.HasValue)
            {
                return Ok(new
                {
                    JobId = existingJobId.Value,
                    Status = "Duplicate",
                    Message = "This file was already accepted; returning the original job."
                });
            }

            var parseResult = await _excelParser.ParseStreamAsync(new MemoryStream(fileBytes), file.FileName, cancellationToken);

            if (parseResult.TotalRowsProcessed > _thresholds.MaxExcelRows)
            {
                return BadRequest(new ProblemDetails
                {
                    Status = StatusCodes.Status400BadRequest,
                    Title = "Row Threshold Exceeded",
                    Detail = $"File contains {parseResult.TotalRowsProcessed} rows, exceeding maximum limit of {_thresholds.MaxExcelRows} rows."
                });
            }

            var jobId = Guid.NewGuid();
            var job = new MigrationJob
            {
                JobId = jobId,
                SourceChannel = "EXCEL_UPLOAD",
                Items = parseResult.Items,
                NotificationEmail = notificationEmail,
                CreatedAtUtc = DateTime.UtcNow
            };

            // Persist the raw uploaded file to storage/file for traceability and reprocessing.
            await _fileStore.SaveAsync(jobId, file.FileName, fileBytes, cancellationToken);

            await _queue.EnqueueAsync(job, cancellationToken);
            await _idempotency.SaveAsync(payloadKey, jobId, cancellationToken);

            return Accepted(new
            {
                JobId = jobId,
                Status = "Queued",
                Format = parseResult.DetectedFormat,
                RowsAccepted = parseResult.Items.Count,
                DuplicateRows = parseResult.DuplicateRows,
                RowErrors = parseResult.RowErrors,
                Message = $"File accepted. A summary report will be emailed to '{notificationEmail}' once complete."
            });
        }

        private static List<ContentItem> DeduplicateByExternalId(IEnumerable<ContentItem> items)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<ContentItem>();

            foreach (var item in items)
            {
                if (!string.IsNullOrWhiteSpace(item.ExternalId) && !seen.Add(item.ExternalId))
                    continue;
                result.Add(item);
            }

            return result;
        }

        private static string ComputeHash(byte[] payload)
        {
            using var sha = SHA256.Create();
            return Convert.ToHexString(sha.ComputeHash(payload));
        }
    }
}
