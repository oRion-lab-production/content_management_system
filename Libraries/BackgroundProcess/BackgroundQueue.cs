using Libraries.LibModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Channels;
using System.Threading.Tasks;

namespace Libraries.BackgroundProcess
{
    public class BackgroundQueue : IBackgroundQueue
    {
        private readonly Channel<MigrationJob> _channel;

        public BackgroundQueue()
        {
            var options = new UnboundedChannelOptions { SingleReader = true };
            _channel = Channel.CreateUnbounded<MigrationJob>(options);
        }

        public async ValueTask EnqueueAsync(MigrationJob job, CancellationToken cancellationToken = default)
            => await _channel.Writer.WriteAsync(job, cancellationToken);

        public async ValueTask<MigrationJob> DequeueAsync(CancellationToken cancellationToken = default)
            => await _channel.Reader.ReadAsync(cancellationToken);
    }
}
