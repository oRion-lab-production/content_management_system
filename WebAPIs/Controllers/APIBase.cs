using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Routing;

namespace WebAPIs.Controllers
{
    [ApiController]
    [Route("/api/v1/[controller]/[action]", Name = "BaseApiRoute_[controller]_[action]")]
    public abstract class APIBase : ControllerBase
    {
        protected APIBase() { }
    }
}
