using Libraries.LibModels;
using Libraries.Mailing;
using Libraries.NoSqlDocumentRepositorys;
using Libraries.WebhookNotifier;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.BackgroundProcess.Workers
{
    public class MigrationWorker : BackgroundService
    {
        private readonly IBackgroundQueue _queue;
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly ILogger<MigrationWorker> _logger;

        public MigrationWorker(
            IBackgroundQueue queue,
            IServiceScopeFactory scopeFactory,
            ILogger<MigrationWorker> logger)
        {
            _queue = queue;
            _scopeFactory = scopeFactory;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("WCMS Background Migration Worker started with Scoped Service Factory.");

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    var job = await _queue.DequeueAsync(stoppingToken);

                    _logger.LogInformation("Processing migration job {JobId} from source: {Channel}", job.JobId, job.SourceChannel);

                    using (var scope = _scopeFactory.CreateScope())
                    {
                        var noSqlRepository = scope.ServiceProvider.GetRequiredService<INoSqlDocumentRepository>();
                        var emailService = scope.ServiceProvider.GetRequiredService<IEmailService>();
                        var upstreamNotifier = scope.ServiceProvider.GetRequiredService<IUpstreamNotifier>();

                        var saveResult = await noSqlRepository.SaveDocumentsAsync(job.JobId, job.Items, job.SourceChannel, stoppingToken);
                        int savedCount = saveResult.ImportedCount;

                        if (job.SourceChannel == "EXCEL_UPLOAD" && !string.IsNullOrEmpty(job.NotificationEmail))
                        {
                            var reportErrors = new List<string>();
                            if (saveResult.SkippedCount > 0)
                                reportErrors.Add($"{saveResult.SkippedCount} duplicate item(s) were skipped (already imported).");

                            await emailService.SendCompletionEmailAsync(
                                job.NotificationEmail,
                                job.JobId,
                                total: job.Items.Count,
                                success: savedCount,
                                errors: reportErrors,
                                stoppingToken
                            );
                        }

                        var importEvent = new ContentImportedEvent
                        {
                            JobId = job.JobId,
                            SourceChannel = job.SourceChannel,
                            ItemsImportedCount = savedCount,
                            ItemsSkippedCount = saveResult.SkippedCount,
                            CompletedAtUtc = DateTime.UtcNow
                        };

                        await upstreamNotifier.NotifyUpstreamSystemsAsync(importEvent, stoppingToken);
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing job in background worker.");
                }
            }
        }
    }
}
