using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using System.Threading.Tasks;

class GraphHelper
{
    private static Settings? _settings;
    private static ClientSecretCredential? _clientSecretCredential;
    private static GraphServiceClient? _appClient;

    public static void InitializeGraphForAppOnlyAuth(Settings settings)
    {
        _settings = settings ?? throw new System.NullReferenceException("Settings cannot be null");

        if (_clientSecretCredential == null)
        {
            _clientSecretCredential = new ClientSecretCredential(
                _settings.TenantId, _settings.ClientId, _settings.ClientSecret);
        }

        if (_appClient == null)
        {
            _appClient = new GraphServiceClient(_clientSecretCredential,
                new[] { "https://graph.microsoft.com/.default" });

        }
    }

    public static async Task<string> GetAppOnlyTokenAsync()
    {
        _ = _clientSecretCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var context = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
        var response = await _clientSecretCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<UserCollectionResponse?> GetUsersAsync()
    {
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Users.GetAsync((config) =>
        {
            config.QueryParameters.Select = new[] { "displayName", "id", "mail" };
            config.QueryParameters.Top = 25;
            config.QueryParameters.Orderby = new[] { "displayName" };
        });
    }

    public static async Task<string> GetReportAsync()
{
    _ = _appClient ??
        throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

    var result = await _appClient.DeviceManagement.DeviceCompliancePolicyDeviceStateSummary.GetAsync();
    string json = JsonConvert.SerializeObject(result, Formatting.Indented);
    return json;
}

    public static async Task<string> GetTroubleshootingEventsAsync()
{
    _ = _appClient ??
        throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

    var result = await _appClient.DeviceManagement.TroubleshootingEvents.GetAsync();
    string json = JsonConvert.SerializeObject(result, Formatting.Indented);
    return json;
}

public static async Task<string> GetFailedAutopilotEventsAsync()
{
    _ = _appClient ?? throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

    var httpClient = new HttpClient();
    var request = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/beta/deviceManagement/autopilotEvents?$filter=deploymentState eq 'failure'");

    // Get the access token
    var tokenRequestContext = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
    var accessToken = await _clientSecretCredential.GetTokenAsync(tokenRequestContext);

    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken.Token);

    var response = await httpClient.SendAsync(request);
    var result = await response.Content.ReadAsStringAsync();

    var formattedResult = JsonConvert.SerializeObject(JsonConvert.DeserializeObject(result), Formatting.Indented);

    return formattedResult;
}
}
