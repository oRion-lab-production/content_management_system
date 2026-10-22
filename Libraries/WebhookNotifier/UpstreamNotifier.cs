using Libraries.LibModels;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.WebhookNotifier
{
    public class UpstreamNotifier : IUpstreamNotifier
    {
        private readonly HttpClient _httpClient;
        private readonly IConfiguration _config;
        private readonly ILogger<UpstreamNotifier> _logger;

        public UpstreamNotifier(HttpClient httpClient, IConfiguration config, ILogger<UpstreamNotifier> logger)
        {
            _httpClient = httpClient;
            _config = config;
            _logger = logger;
        }

        public async Task NotifyUpstreamSystemsAsync(ContentImportedEvent importEvent, CancellationToken cancellationToken)
        {
            var searchWebhook = _config["UpstreamSettings:SearchEngineWebhookUrl"]!;
            var mobileWebhook = _config["UpstreamSettings:MobileAppWebhookUrl"]!;

            await SendWebhookAsync(searchWebhook, importEvent, "Search Engine", cancellationToken);
            await SendWebhookAsync(mobileWebhook, importEvent, "Mobile App Backend", cancellationToken);
        }

        private async Task SendWebhookAsync(string url, ContentImportedEvent payload, string systemName, CancellationToken cancellationToken)
        {
            try
            {
                _logger.LogInformation("[UPSTREAM NOTIFIER] Sending webhook signal to {System} at {Url}", systemName, url);
                var response = await _httpClient.PostAsJsonAsync(url, payload, cancellationToken);
                if (response.IsSuccessStatusCode)
                {
                    _logger.LogInformation("[UPSTREAM NOTIFIER] Successfully notified {System}.", systemName);
                }
                else
                {
                    _logger.LogWarning("[UPSTREAM NOTIFIER] {System} returned non-success code: {Status}", systemName, response.StatusCode);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[UPSTREAM NOTIFIER] Failed to reach {System} endpoint.", systemName);
            }
        }
    }
}
