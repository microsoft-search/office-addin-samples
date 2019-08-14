using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;
using GraphSearchApi.Extensions;
using GraphSearchApi.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Client;

namespace GraphSearchApi
{
    [Microsoft.AspNetCore.Mvc.Route("api/[controller]")]
    [ApiController]
    public class GraphController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly IGraphSdkHelper _graphSdkHelper;
        private readonly IGraphAuthProvider _authProvider;

        public GraphController(IConfiguration configuration, IGraphSdkHelper graphSdkHelper, IGraphAuthProvider authProvider)
        {
            _configuration = configuration;
            _graphSdkHelper = graphSdkHelper;
            _authProvider = authProvider;
        }

        [Microsoft.AspNetCore.Mvc.HttpPost]
        [Microsoft.AspNetCore.Mvc.Route("Token")]
        public string Token()
        {
            var userIdentity = (ClaimsIdentity)User.Identity;
            var identifier = userIdentity.FindFirst(Startup.ObjectIdentifierType)?.Value + "." + userIdentity.FindFirst(Startup.TenantIdType)?.Value;
            var accessToken = _authProvider.GetUserAccessTokenAsync(identifier).Result;
            return accessToken;
        }
    }
}
