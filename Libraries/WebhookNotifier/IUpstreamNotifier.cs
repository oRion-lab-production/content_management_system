using Libraries.LibModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.WebhookNotifier
{
    public interface IUpstreamNotifier
    {
        Task NotifyUpstreamSystemsAsync(ContentImportedEvent importEvent, CancellationToken cancellationToken);
    }
}
