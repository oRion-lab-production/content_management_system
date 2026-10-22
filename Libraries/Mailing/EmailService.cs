using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.Mailing
{
    public class EmailService : IEmailService
    {
        private readonly ILogger<EmailService> _logger;

        public EmailService(ILogger<EmailService> logger) => _logger = logger;

        public Task SendCompletionEmailAsync(string email, Guid jobId, int total, int success, List<string> errors, CancellationToken cancellationToken)
        {
            _logger.LogInformation("==================================================");
            _logger.LogInformation("[EMAIL DISPATCHED] To: {Email}", email);
            _logger.LogInformation("Subject: Excel Migration Complete - Job #{JobId}", jobId);
            _logger.LogInformation("Summary: Total Rows: {Total} | Imported: {Success} | Failed: {Failed}", total, success, errors.Count);
            if (errors.Any())
            {
                _logger.LogInformation("Row Errors:\n - {Errors}", string.Join("\n - ", errors));
            }
            _logger.LogInformation("==================================================");
            return Task.CompletedTask;
        }
    }
}
