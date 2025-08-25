using Azure.Core;
using Microsoft.Graph;
using System.IdentityModel.Tokens.Jwt;

namespace O365C.Agent.ProjectAssist.Bot.Helpers
{
    public class GraphAuthHelper
    {
        public static GraphServiceClient CreateGraphClientWithAccessToken(string accessToken)
        {
            var expiresOn = GetExpirationFromAccessToken(accessToken);

            // Ensure expiresOn is not null before passing it to the constructor
            if (!expiresOn.HasValue)
            {
                throw new ArgumentException("The access token does not contain a valid expiration time.", nameof(accessToken));
            }

            // Create a custom TokenCredential that returns the provided access token.
            var tokenCredential = new StaticAccessTokenCredential(accessToken, expiresOn.Value);

            // Create the GraphServiceClient using the custom TokenCredential.
            // The GraphServiceClient will use this credential to add the Authorization header
            // with the bearer token to all requests.
            var graphClient = new GraphServiceClient(tokenCredential);

            return graphClient;
        }

        private class StaticAccessTokenCredential : TokenCredential
        {
            private readonly string _accessToken;
            private readonly DateTimeOffset _expiresOn;

            public StaticAccessTokenCredential(string accessToken, DateTimeOffset expiresOn)
            {
                _accessToken = accessToken ?? throw new ArgumentNullException(nameof(accessToken));
                _expiresOn = expiresOn;
            }

            public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
            {
                return new AccessToken(_accessToken, _expiresOn);
            }

            public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
            {
                return new ValueTask<AccessToken>(GetToken(requestContext, cancellationToken));
            }
        }

        public static DateTimeOffset? GetExpirationFromAccessToken(string accessToken)
        {
            if (string.IsNullOrWhiteSpace(accessToken))
            {
                Console.WriteLine("Access token is null or empty.");
                return null;
            }

            try
            {
                var handler = new JwtSecurityTokenHandler();

                if (!handler.CanReadToken(accessToken))
                {
                    Console.WriteLine("Access token is not a valid JWT format.");
                    return null;
                }

                var jwtToken = handler.ReadJwtToken(accessToken);

                return jwtToken.ValidTo;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error decoding access token: {ex.Message}");
                return null;
            }
        }


        /// <summary>
        /// Creates a GraphServiceClient using client credentials (client secret) flow.
        /// </summary>
        /// <param name="tenantId">Azure AD tenant ID</param>
        /// <param name="clientId">Azure AD application (client) ID</param>
        /// <param name="clientSecret">Azure AD application client secret</param>
        /// <param name="scopes">Scopes to request (should be new[] { "https://graph.microsoft.com/.default" })</param>
        /// <returns>GraphServiceClient instance</returns>
        public static GraphServiceClient CreateGraphClientWithClientSecret(
            string tenantId,
            string clientId,
            string clientSecret,
            string[] scopes)
        {
            if (string.IsNullOrWhiteSpace(tenantId))
                throw new ArgumentException("Tenant ID cannot be null or empty.", nameof(tenantId));
            if (string.IsNullOrWhiteSpace(clientId))
                throw new ArgumentException("Client ID cannot be null or empty.", nameof(clientId));
            if (string.IsNullOrWhiteSpace(clientSecret))
                throw new ArgumentException("Client Secret cannot be null or empty.", nameof(clientSecret));
            if (scopes == null || scopes.Length == 0)
                throw new ArgumentException("Scopes cannot be null or empty.", nameof(scopes));

            var options = new Azure.Identity.ClientSecretCredentialOptions
            {
                AuthorityHost = Azure.Identity.AzureAuthorityHosts.AzurePublicCloud,
            };

            var clientSecretCredential = new Azure.Identity.ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
            return graphClient;
        }
    }
}
