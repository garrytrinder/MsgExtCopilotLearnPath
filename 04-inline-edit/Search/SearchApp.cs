using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using AdaptiveCards;
using Newtonsoft.Json.Linq;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Graph.Models;
using AdaptiveCards.Templating;
using Newtonsoft.Json;
using MsgExtProductSupport.Models;

namespace MsgExtProductSupport.Search;

public class SearchApp : TeamsActivityHandler
{
    private readonly string connectionName;
    private readonly string spoHostname;
    private readonly string spoSiteUrl;

    public SearchApp(IConfiguration configuration)
    {
        connectionName = configuration["CONNECTION_NAME"];
        spoHostname = configuration["SPO_HOSTNAME"];
        spoSiteUrl = configuration["SPO_SITE_URL"];
    }

    protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
    {
        var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
        var tokenResponse = await GetToken(userTokenClient, query.State, turnContext.Activity.From.Id, turnContext.Activity.ChannelId, connectionName, cancellationToken);

        if (!HasToken(tokenResponse))
        {
            return await CreateAuthResponse(userTokenClient, connectionName, (Activity)turnContext.Activity, cancellationToken);
        }

        var name = GetQueryData(query.Parameters, "ProductName");
        var nameFilter = !string.IsNullOrEmpty(name) ? $"startswith(fields/Title, '{name}')" : string.Empty;
        var filters = new List<string> { nameFilter };
        var filterQuery = filters.Count == 1 ? filters.FirstOrDefault() : string.Join(" and ", filters);

        var graphClient = CreateGraphClient(tokenResponse);
        var site = await GetProductMarketingSite(graphClient, spoHostname, spoSiteUrl, cancellationToken);
        var items = await GetProducts(graphClient, site.SharepointIds.SiteId, filterQuery, cancellationToken);
        var drive = await GetSharePointDrive(graphClient, site.SharepointIds.SiteId, "Product Imagery", cancellationToken);

        var adaptiveCardJson = File.ReadAllText(@"AdaptiveCards\Product.json");
        var template = new AdaptiveCardTemplate(adaptiveCardJson);

        var attachments = new List<MessagingExtensionAttachment>();

        foreach (var item in items.Value)
        {

            var product = JsonConvert.DeserializeObject<Product>(item.AdditionalData["fields"].ToString());
            product.Id = item.Id;

            var thumbnails = await GetThumbnails(graphClient, drive.Id, product.PhotoSubmission, cancellationToken);

            var resultCard = template.Expand(new
            {
                Product = product,
                ProductImage = thumbnails.Large.Url,
                SPOHostname = spoHostname,
                SPOSiteUrl = spoSiteUrl,
            });

            var previewcard = new ThumbnailCard
            {
                Title = product.Title,
                Subtitle = product.RetailCategory,
                Images = new List<CardImage> { new() { Url = thumbnails.Small.Url } }
            }.ToAttachment();

            var attachment = new MessagingExtensionAttachment
            {
                Content = JsonConvert.DeserializeObject(resultCard),
                ContentType = AdaptiveCard.ContentType,
                Preview = previewcard
            };

            attachments.Add(attachment);
        }

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

    protected override async Task<AdaptiveCardInvokeResponse> OnAdaptiveCardInvokeAsync(ITurnContext<IInvokeActivity> turnContext, AdaptiveCardInvokeValue invokeValue, CancellationToken cancellationToken)
    {
        var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
        var tokenResponse = await GetToken(userTokenClient, invokeValue.State, turnContext.Activity.From.Id, turnContext.Activity.ChannelId, connectionName, cancellationToken);

        if (!HasToken(tokenResponse))
        {
            return await CreateOAuthCardResponse(userTokenClient, connectionName, (Activity)turnContext.Activity, cancellationToken);
        }

        var data = JsonConvert.DeserializeObject<ActionData>(invokeValue.Action.Data.ToString());

        var graphClient = CreateGraphClient(tokenResponse);
        var site = await GetProductMarketingSite(graphClient, spoHostname, spoSiteUrl, cancellationToken);
        var product = await GetProduct(data.Id, graphClient, site.SharepointIds.SiteId, cancellationToken);
        var drive = await GetSharePointDrive(graphClient, site.SharepointIds.SiteId, "Product Imagery", cancellationToken);

        return invokeValue.Action.Verb switch
        {
            "edit" => await HandleEditAction(graphClient, site, product, cancellationToken),
            "edit-save" => await HandleSaveAction(data, graphClient, site.SharepointIds.SiteId, data.Id, product, drive.Id, cancellationToken),
            "edit-cancel" => await HandleCancelAction(graphClient, product, drive.Id, cancellationToken),
            _ => null
        };
    }

    private async Task<AdaptiveCardInvokeResponse> HandleCancelAction(GraphServiceClient graphClient, Product product, string driveId, CancellationToken cancellationToken)
    {
        ThumbnailSet thumbnails = await GetThumbnails(graphClient, driveId, product.PhotoSubmission, cancellationToken);

        return new AdaptiveCardInvokeResponse
        {
            StatusCode = StatusCodes.Status200OK,
            Type = AdaptiveCard.ContentType,
            Value = CreateAdaptiveCard(@"AdaptiveCards\Product.json", new
            {
                Product = product,
                ProductImage = thumbnails.Large.Url,
                SPOHostname = spoHostname,
                SPOSiteUrl = spoSiteUrl
            })
        };
    }

    private static async Task<Product> UpdateProduct(string itemId, GraphServiceClient graphClient, string siteId, ActionData data, CancellationToken cancellationToken)
    {
        var fields = await graphClient.Sites[siteId].Lists["Products"].Items[itemId].Fields.PatchAsync(new FieldValueSet
        {
            AdditionalData = new Dictionary<string, object>
            {
                { "Title", data.Title },
                { "RetailCategory", data.RetailCategory },
                { "ReleaseDate", data.ReleaseDate.ToString("u") }
            }
        }, cancellationToken: cancellationToken);

        fields.AdditionalData["Id"] = itemId;

        var json = JsonConvert.SerializeObject(fields.AdditionalData);
        return JsonConvert.DeserializeObject<Product>(json);
    }

    private async Task<AdaptiveCardInvokeResponse> HandleSaveAction(ActionData data, GraphServiceClient graphClient, string siteId, string itemId, Product product, string driveId, CancellationToken cancellationToken)
    {
        var updatedProduct = await UpdateProduct(itemId, graphClient, siteId, data, cancellationToken);
        ThumbnailSet thumbnails = await GetThumbnails(graphClient, driveId, product.PhotoSubmission, cancellationToken);

        return new AdaptiveCardInvokeResponse
        {
            StatusCode = StatusCodes.Status200OK,
            Type = AdaptiveCard.ContentType,
            Value = CreateAdaptiveCard(@"AdaptiveCards\Product.json", new
            {
                Product = updatedProduct,
                ProductImage = thumbnails.Large.Url,
                SPOHostname = spoHostname,
                SPOSiteUrl = spoSiteUrl
            })
        };
    }

    private static async Task<AdaptiveCardInvokeResponse> HandleEditAction(GraphServiceClient graphClient, Site site, Product product, CancellationToken cancellationToken)
    {
        var retailCategories = await GetRetailCategories(graphClient, site.SharepointIds.SiteId, cancellationToken);

        return new AdaptiveCardInvokeResponse
        {
            StatusCode = StatusCodes.Status200OK,
            Type = AdaptiveCard.ContentType,
            Value = CreateAdaptiveCard(
                @"AdaptiveCards\EditForm.json",
                new
                {
                    Product = product,
                    RetailCategories = retailCategories
                })
        };
    }

    private static object CreateAdaptiveCard(string path, object data)
    {
        var templateText = File.ReadAllText(path);
        var template = new AdaptiveCardTemplate(templateText);
        var result = template.Expand(data);

        return JsonConvert.DeserializeObject(result);
    }

    private static async Task<List<string>> GetRetailCategories(GraphServiceClient graphClient, string siteId, CancellationToken cancellationToken)
    {
        var column = await graphClient.Sites[siteId].Lists["Products"].Columns["RetailCategory"].GetAsync(null, cancellationToken);
        return column.Choice.Choices;
    }

    private static async Task<Product> GetProduct(string itemId, GraphServiceClient graphClient, string siteId, CancellationToken cancellationToken)
    {
        var item = await graphClient.Sites[siteId].Lists["Products"].Items[itemId].GetAsync(r => r.QueryParameters.Expand = new string[] { "fields" }, cancellationToken);
        item.Fields.AdditionalData["Id"] = itemId;
        var json = JsonConvert.SerializeObject(item.Fields.AdditionalData);
        return JsonConvert.DeserializeObject<Product>(json);
    }

    private static async Task<AdaptiveCardInvokeResponse> CreateOAuthCardResponse(UserTokenClient userTokenClient, string connectionName, Activity activity, CancellationToken cancellationToken)
    {
        var resource = await userTokenClient.GetSignInResourceAsync(connectionName, activity, null, cancellationToken);

        return new AdaptiveCardInvokeResponse
        {
            StatusCode = 401,
            Type = $"{Activity.ContentType}.loginRequest",
            Value = JObject.FromObject(new OAuthCard
            {
                Buttons = new List<CardAction>
                {
                    new() {
                        Title = "Sign In",
                        Type = ActionTypes.Signin,
                        Value = resource.SignInLink
                    }
                },
                Text = "Please sign in to continue.",
                ConnectionName = connectionName,
            })
        };
    }

    private static async Task<ThumbnailSet> GetThumbnails(GraphServiceClient graphClient, string driveId, string photoUrl, CancellationToken cancellationToken)
    {
        var fileName = photoUrl.Split('/').Last();
        var driveItem = await graphClient.Drives[driveId].Root.ItemWithPath(fileName).GetAsync(null, cancellationToken);
        var thumbnails = await graphClient.Drives[driveId].Items[driveItem.Id].Thumbnails["0"].GetAsync(r => r.QueryParameters.Select = new string[] { "small", "large" }, cancellationToken);
        return thumbnails;
    }

    private static async Task<Drive> GetSharePointDrive(GraphServiceClient graphClient, string siteId, string name, CancellationToken cancellationToken)
    {
        var drives = await graphClient.Sites[siteId].Drives.GetAsync(r => r.QueryParameters.Select = new string[] { "id", "name" }, cancellationToken);
        var drive = drives.Value.Find(d => d.Name == name);
        return drive;
    }

    private static async Task<SiteCollectionResponse> GetProducts(GraphServiceClient graphClient, string siteId, string filterQuery, CancellationToken cancellationToken)
    {
        var fields = new string[]
        {
            "fields/Id",
            "fields/Title",
            "fields/RetailCategory",
            "fields/PhotoSubmission",
            "fields/CustomerRating",
            "fields/ReleaseDate"
        };

        var requestUrl = string.IsNullOrEmpty(filterQuery)
            ? $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/Products/items?expand={string.Join(",", fields)}"
            : $"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/Products/items?expand={string.Join(",", fields)}&$filter={filterQuery}";

        var request = graphClient.Sites.WithUrl(requestUrl);
        return await request.GetAsync(null, cancellationToken);
    }

    private static async Task<Site> GetProductMarketingSite(GraphServiceClient graphClient, string hostName, string siteUrl, CancellationToken cancellationToken)
    {
        return await graphClient.Sites[$"{hostName}:/{siteUrl}"].GetAsync(r => r.QueryParameters.Select = new string[] { "sharePointIds" }, cancellationToken);
    }

    private static GraphServiceClient CreateGraphClient(TokenResponse tokenResponse)
    {
        TokenProvider provider = new() { Token = tokenResponse.Token };
        var authenticationProvider = new BaseBearerTokenAuthenticationProvider(provider);
        var graphClient = new GraphServiceClient(authenticationProvider);
        return graphClient;
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

    private static string GetQueryData(IList<MessagingExtensionParameter> parameters, string key)
    {
        if (parameters.Any() != true)
        {
            return string.Empty;
        }

        var foundPair = parameters.FirstOrDefault(pair => pair.Name == key);
        return foundPair?.Value?.ToString() ?? string.Empty;
    }
}