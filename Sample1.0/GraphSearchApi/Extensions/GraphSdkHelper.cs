using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;

namespace GraphSearchApi.Extensions
{
    public class GraphSdkHelper : IGraphSdkHelper
    {
        private readonly IGraphAuthProvider _authProvider;
        private GraphServiceClient _graphClient;
        
        public GraphSdkHelper(IGraphAuthProvider authProvider)
        {
            _authProvider = authProvider;
        }

        // Get an authenticated Microsoft Graph Service client.
        public GraphServiceClient GetAuthenticatedClient(ClaimsIdentity userIdentity)
        {
            _graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                async requestMessage =>
                {
                    // Get user's id for token cache.
                    var identifier = userIdentity.FindFirst(Startup.ObjectIdentifierType)?.Value + "." + userIdentity.FindFirst(Startup.TenantIdType)?.Value;

                    // Passing tenant ID to the sample auth provider to use as a cache key
                    var accessToken = await _authProvider.GetUserAccessTokenAsync(identifier);

                    // Append the access token to the request
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));

            return _graphClient;
        }
    }
    public interface IGraphSdkHelper
    {
        GraphServiceClient GetAuthenticatedClient(ClaimsIdentity userIdentity);
    }
}
