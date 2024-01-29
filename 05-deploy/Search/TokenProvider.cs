using Microsoft.Kiota.Abstractions.Authentication;

namespace MsgExtProductSupport.Search
{
    public class TokenProvider : IAccessTokenProvider
    {
        public string Token { get; set; }
        public AllowedHostsValidator AllowedHostsValidator => throw new NotImplementedException();

        public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(Token);
        }
    }
}
