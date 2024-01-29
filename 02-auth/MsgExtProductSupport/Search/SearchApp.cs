using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using AdaptiveCards;
using Newtonsoft.Json.Linq;
using Microsoft.Bot.Connector.Authentication;

namespace MsgExtProductSupport.Search;

public class SearchApp : TeamsActivityHandler
{
    private readonly string connectionName;

    public SearchApp(IConfiguration configuration)
    {
        connectionName = configuration["CONNECTION_NAME"];
    }

    protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
    {
        var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
        var tokenResponse = await GetToken(userTokenClient, query.State, turnContext.Activity.From.Id, turnContext.Activity.ChannelId, connectionName, cancellationToken);

        if (!HasToken(tokenResponse))
        {
            return await CreateAuthResponse(userTokenClient, connectionName, (Activity)turnContext.Activity, cancellationToken);
        }

        var text = query?.Parameters?[0]?.Value as string ?? string.Empty;
        var packages = await FindPackages(text);
        // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
        var attachments = packages.Select(package =>
        {
            var previewCard = new ThumbnailCard { Title = package.Item1 };
            var adaptiveCard = new AdaptiveCard(new AdaptiveSchemaVersion("1.4"))
            {
                Body = new List<AdaptiveElement>
            {
                new AdaptiveTextBlock
                {
                    Text = package.Item1,
                    Size = AdaptiveTextSize.Large,
                    Wrap = true
                },
                new AdaptiveTextBlock
                {
                    Text = package.Item3,
                    Size = AdaptiveTextSize.Medium,
                    Wrap = true
                }
            }
            };
            if (!string.IsNullOrEmpty(package.Item5))
            {
                previewCard.Images = new List<CardImage>() { new CardImage(package.Item5, "Icon") };
                adaptiveCard.Body.Insert(0, new AdaptiveImage()
                {
                    Url = new Uri(package.Item5),
                    Style = AdaptiveImageStyle.Person,
                    Size = AdaptiveImageSize.Small,
                });
            }
            var attachment = new MessagingExtensionAttachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
                Preview = previewCard.ToAttachment()
            };

            return attachment;
        }).ToList();

        return new MessagingExtensionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = "list",
                Attachments = attachments
            }
        };
    }

    // Generate a set of substrings to illustrate the idea of a set of results coming back from a query. 
    private async Task<IEnumerable<(string, string, string, string, string)>> FindPackages(string text)
    {
        var httpClient = new HttpClient();
        var response = await httpClient.GetStringAsync($"https://azuresearch-usnc.nuget.org/query?q=id:{text}&prerelease=true");
        var obj = JObject.Parse(response);
        return obj["data"].Select(item => (
            item["id"].ToString(),
            item["version"].ToString(),
            item["description"].ToString(),
            item["projectUrl"]?.ToString(),
            item["iconUrl"]?.ToString()));
    }
	
    private static async Task<TokenResponse> GetToken(UserTokenClient userTokenClient, string state, string userId, string channelId, string connectionName, CancellationToken cancellationToken)
    {
        var magicCode = string.Empty;

        if (!string.IsNullOrEmpty(state))
        {
            if (int.TryParse(state, out var parsed))
            {
                magicCode = parsed.ToString();
            }
        }

        return await userTokenClient.GetUserTokenAsync(userId, connectionName, channelId, magicCode, cancellationToken);
    }
	
    private static bool HasToken(TokenResponse tokenResponse)
    {
        return tokenResponse != null && !string.IsNullOrEmpty(tokenResponse.Token);
    }

    private static async Task<MessagingExtensionResponse> CreateAuthResponse(UserTokenClient userTokenClient, string connectionName, Activity activity, CancellationToken cancellationToken)
    {
        var resource = await userTokenClient.GetSignInResourceAsync(connectionName, activity, null, cancellationToken);

        return new MessagingExtensionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "auth",
                SuggestedActions = new MessagingExtensionSuggestedAction
                {
                    Actions = new List<CardAction>
                    {
                        new() {
                            Type = ActionTypes.OpenUrl,
                            Value = resource.SignInLink,
                            Title = "Sign In",
                        },
                    },
                },
            },
        };
    }
}