using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace GraphSearchApiExcelWeb
{
    public class GraphController : ApiController
    {
        [System.Web.Http.HttpPost]
        public HttpResponseMessage Token(Token body)
        {
            if (string.IsNullOrEmpty(body.UserToken))
            {
                return SendErrorToClient(HttpStatusCode.InternalServerError, null, "No token provided");
            }

            UserAssertion userAssertion = new UserAssertion(body.UserToken);

            // Get the access token for MS Graph. 
            ConfidentialClientApplicationBuilder b = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"]);
            b.WithClientSecret(ConfigurationManager.AppSettings["ida:Password"]).WithRedirectUri(ConfigurationManager.AppSettings["ida:RedirectUri"]);
            IConfidentialClientApplication cca = b.Build();

            //string[] graphScopes = { "user.read", "files.read.all", "mail.read", "calendars.read" };

            string[] graphScopes = { "Mail.Read", "Files.Read.All", "Calendars.Read" };

            AuthenticationResult result = null;
            try
            {
                // The AcquireTokenOnBehalfOfAsync method will first look in the MSAL in memory cache for a
                // matching access token. Only if there isn't one, does it initiate the "on behalf of" flow
                // with the Azure AD V2 endpoint.
                result = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion).ExecuteAsync().Result;
            }
            catch (MsalServiceException e)
            {
                if (e.Message.StartsWith("AADSTS50076"))
                {

                    string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
                    return SendErrorToClient(HttpStatusCode.InternalServerError, e, responseMessage);
                }

                if ((e.Message.StartsWith("AADSTS65001"))
                || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
                {
                    return SendErrorToClient(HttpStatusCode.InternalServerError, e, e.Message);
                }

                return SendErrorToClient(HttpStatusCode.InternalServerError, e, e.Message);
            }

            HttpRequestMessage requestMessage = new HttpRequestMessage();
            var configuration = new HttpConfiguration();
            requestMessage.Properties[System.Web.Http.Hosting.HttpPropertyKeys.HttpConfigurationKey] = configuration;
            HttpResponseMessage tokenMessage = requestMessage.CreateResponse(HttpStatusCode.OK, result.AccessToken);

            return tokenMessage;
        }

        private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
        {
            HttpError error;

            if (e != null)
            {
                error = new HttpError(e, true);
            }
            else
            {
                error = new HttpError(message);
            }

            HttpRequestMessage requestMessage = new HttpRequestMessage();
            HttpResponseMessage errorMessage = requestMessage.CreateErrorResponse(statusCode, error);

            return errorMessage;
        }
    }
}