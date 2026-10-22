using Microsoft.AspNetCore.Mvc;

namespace WebAPIs.Controllers
{
    [ApiController]
    [Route("/api/v2/[controller]/[action]", Name = "BaseApiRoute_[controller]_[action]")]
    public abstract class APIBase_V2 : ControllerBase
    {
        protected APIBase_V2() { }
    }
}
