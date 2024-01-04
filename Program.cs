Console.WriteLine(".NET Graph App-only Tutorial\n");

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. List users");
    Console.WriteLine("3. Get Compliance Report");
    Console.WriteLine("4. Get Troubleshooting Events");
    Console.WriteLine("5. Get Failed Autopilot Events");
    Console.WriteLine("6. Get Device Configuration Conflict Summary");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            // List users
            await ListUsersAsync();
            break;
        case 3:
            // Run any Graph code
            await MakeGraphCallAsync();
            break;
        case 4:
            // Run any Graph code
            await GetTroubleshootingEventsAsync();
            break;  
        case 5:
            // Run any Graph code
            await GetFailedAutopilotEventsAsync();
            break;  
        case 6:
            // Run any Graph code
            await GetDeviceConfigurationConflictSummaryAsync();
            break;    
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}
void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForAppOnlyAuth(settings);
}

async Task DisplayAccessTokenAsync()
{
    try
    {
        var appOnlyToken = await GraphHelper.GetAppOnlyTokenAsync();
        Console.WriteLine($"App-only token: {appOnlyToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting app-only access token: {ex.Message}");
    }
}

async Task ListUsersAsync()
{
    try
    {
        var userPage = await GraphHelper.GetUsersAsync();

        if (userPage?.Value == null)
        {
            Console.WriteLine("No results returned.");
            return;
        }

        // Output each users's details
        foreach (var user in userPage.Value)
        {
            Console.WriteLine($"User: {user.DisplayName ?? "NO NAME"}");
            Console.WriteLine($"  ID: {user.Id}");
            Console.WriteLine($"  Email: {user.Mail ?? "NO EMAIL"}");
        }

        // If NextPageRequest is not null, there are more users
        // available on the server
        // Access the next page like:
        // var nextPageRequest = new UsersRequestBuilder(userPage.OdataNextLink, _appClient.RequestAdapter);
        // var nextPage = await nextPageRequest.GetAsync();
        var moreAvailable = !string.IsNullOrEmpty(userPage.OdataNextLink);

        Console.WriteLine($"\nMore users available? {moreAvailable}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting users: {ex.Message}");
    }
}

 async Task<string> MakeGraphCallAsync() // Get json report from Graph GetReportAsync
{
    try
    {
    var Result = await GraphHelper.GetReportAsync();
        Console.WriteLine($"Report: {Result}");
        return Result;
        }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting report: {ex.Message}");
        return "error";
    }
  
}
async Task<string> GetTroubleshootingEventsAsync() // Get json report from Graph GetReportAsync
{
    try
    {
    var Result = await GraphHelper.GetTroubleshootingEventsAsync();
        Console.WriteLine($"Report: {Result}");
        return Result;
        }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting report: {ex.Message}");
        return "error";
    }
  
}
async Task<string> GetFailedAutopilotEventsAsync() // Get json report from Graph GetReportAsync
{
    try
    {
    var Result = await GraphHelper.GetFailedAutopilotEventsAsync();
        Console.WriteLine($"Report: {Result}");
        return Result;
        }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting report: {ex.Message}");
        return "error";
    }
  
}
async Task<string> GetDeviceConfigurationConflictSummaryAsync() // Get json report from Graph GetReportAsync
{
    try
    {
    var Result = await GraphHelper.GetDeviceConfigurationConflictSummaryAsync();
        Console.WriteLine($"Report: {Result}");
        return Result;
        }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting report: {ex.Message}");
        return "error";
    }
  
}



