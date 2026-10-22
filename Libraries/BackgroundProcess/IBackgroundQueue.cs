using Libraries.LibModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.BackgroundProcess
{
    public interface IBackgroundQueue
    {
        ValueTask EnqueueAsync(MigrationJob job, CancellationToken cancellationToken = default);
        ValueTask<MigrationJob> DequeueAsync(CancellationToken cancellationToken = default);
    }
}
