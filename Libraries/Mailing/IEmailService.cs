using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Libraries.Mailing
{
    public interface IEmailService
    {
        Task SendCompletionEmailAsync(string email, Guid jobId, int total, int success, List<string> errors, CancellationToken cancellationToken);
    }
}
