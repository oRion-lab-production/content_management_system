using Libraries.BackgroundProcess;
using Libraries.BackgroundProcess.Workers;
using Libraries.ExcelParser;
using Libraries.Idempotency;
using Libraries.Mailing;
using Libraries.NoSqlDocumentRepository;
using Libraries.NoSqlDocumentRepositorys;
using Libraries.Storage;
using Libraries.WebhookNotifier;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries
{
    public static class DependencyIScopeProvider
    {
        public static void ResolveIntegration(this IServiceCollection services)
        {
            services.AddSingleton<IBackgroundQueue, BackgroundQueue>();
            services.AddHostedService<MigrationWorker>();

            services.AddHttpClient();
            services.AddScoped<IUpstreamNotifier, UpstreamNotifier>();
        }

        public static void ResolveServices(this IServiceCollection services)
        {
            services.AddSingleton<IStorageLocator, StorageLocator>();
            services.AddScoped<IExcelParserService, ExcelParserService>();
            services.AddScoped<IUploadFileStore, UploadFileStore>();
            services.AddScoped<IEmailService, EmailService>();
            services.AddSingleton<IIdempotencyStore, FileIdempotencyStore>();
        }

        public static void ResolveRepositories(this IServiceCollection services)
        {
            services.AddScoped<INoSqlDocumentRepository, FileNoSqlDocumentRepository>();
        }
    }
}
