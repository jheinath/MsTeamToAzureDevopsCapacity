using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Identity.Client;

internal class Program
{
    //Make sure to replace the placeholder values <YOUR_CLIENT_ID>, <YOUR_CLIENT_SECRET>,
    //<YOUR_TENANT_ID>, <YOUR_AZURE_DEVOPS_URI>, <YOUR_PAT_TOKEN>,
    //<YOUR_AZURE_DEVOPS_TEAM_ID>, <YOUR_AZURE_DEVOPS_ITERATION_ID>, <YOUR_TEAMS_GROUP_ID>,
    //and <YOUR_TEAMS_CHANNEL_ID> with your actual values.

    private const string ClientId = "<YOUR_CLIENT_ID>"; // Insert your Azure AD application client ID
    private const string ClientSecret = "<YOUR_CLIENT_SECRET>"; // Insert your Azure AD application client secret
    private const string TenantId = "<YOUR_TENANT_ID>"; // Insert your Azure AD tenant ID
    private const string AzureDevOpsUri = "<YOUR_AZURE_DEVOPS_URI>"; // Insert the URI of your Azure DevOps instance (e.g., "https://dev.azure.com/yourorganization")
    private const string PatToken = "<YOUR_PAT_TOKEN>"; // Insert your Azure DevOps personal access token
    private const string TeamId = "<YOUR_AZURE_DEVOPS_TEAM_ID>"; // Insert your Azure DevOps team ID
    private const string IterationId = "<YOUR_AZURE_DEVOPS_ITERATION_ID>"; // Insert the ID of the Azure DevOps iteration
    private const string GraphApiUrl = "https://graph.microsoft.com/v1.0";
    private const string TokenEndpoint = "https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";
    private const string TeamsGroupId = "<YOUR_TEAMS_GROUP_ID>"; // Insert the ID of the Microsoft Teams group
    private const string ChannelId = "<YOUR_TEAMS_CHANNEL_ID>"; // Insert the ID of the Microsoft Teams channel

    private static async Task Main(string[] args)
    {
        // Authenticate and get access token for Microsoft Graph API
        var graphAccessToken = await GetAccessToken();

        // Get start and end dates of the sprint
        var sprintStartDate = await GetSprintStartDate();
        var sprintEndDate = await GetSprintEndDate();

        // Retrieve appointments from the specified Microsoft Teams channel calendar within the sprint range
        var appointments = await GetAppointments(graphAccessToken, sprintStartDate, sprintEndDate);

        // Post appointments to Azure DevOps sprint capacity for each team member
        await PostAppointmentsToSprintCapacity(appointments);

        Console.WriteLine("Appointments posted to Azure DevOps sprint capacity successfully.");
    }

    private static async Task<string> GetAccessToken()
    {
        var clientApp = ConfidentialClientApplicationBuilder.Create(ClientId)
            .WithClientSecret(ClientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{TenantId}/v2.0"))
            .Build();

        var authResult = await clientApp.AcquireTokenForClient(new string[] { $"{GraphApiUrl}/.default" })
            .ExecuteAsync();

        return authResult.AccessToken;
    }

    private static async Task<DateTime> GetSprintStartDate()
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", PatToken);

        const string requestUrl = $"{AzureDevOpsUri}/{TeamId}/_apis/work/teamsettings/iterations/{IterationId}?api-version=6.0";
        var response = await httpClient.GetAsync(requestUrl);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (response.IsSuccessStatusCode)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var iteration = JsonSerializer.Deserialize<AzureDevOpsIteration>(responseContent, options);
            return iteration.Attributes.StartDate;
        }
        else
        {
            throw new Exception($"Failed to retrieve sprint start date from Azure DevOps. Error: {responseContent}");
        }
    }

    private static async Task<DateTime> GetSprintEndDate()
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", PatToken);

        const string requestUrl = $"{AzureDevOpsUri}/{TeamId}/_apis/work/teamsettings/iterations/{IterationId}?api-version=6.0";
        var response = await httpClient.GetAsync(requestUrl);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
            throw new Exception($"Failed to retrieve sprint end date from Azure DevOps. Error: {responseContent}");
        
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        var iteration = JsonSerializer.Deserialize<AzureDevOpsIteration>(responseContent, options);
        return iteration.Attributes.FinishDate;

    }

    private static async Task<List<Appointment>> GetAppointments(string accessToken, DateTime startDate, DateTime endDate)
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        var startDateTime = startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
        var endDateTime = endDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
        var requestUrl = $"{GraphApiUrl}/groups/{TeamsGroupId}/channels/{ChannelId}/calendarView?startDateTime={startDateTime}&endDateTime={endDateTime}";

        var response = await httpClient.GetAsync(requestUrl);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
            throw new Exception($"Failed to retrieve appointments from Microsoft Teams channel calendar. Error: {responseContent}");
        
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        var events = JsonSerializer.Deserialize<GraphCalendarEventsResponse>(responseContent, options);
        return events.Value;

    }

    private static async Task PostAppointmentsToSprintCapacity(List<Appointment> appointments)
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", PatToken);

        var capacityItems = new List<CapacityItem>();
        foreach (var appointment in appointments)
        {
            var capacityItem = new CapacityItem
            {
                TeamId = TeamId,
                IterationId = IterationId,
                TeamMemberEmail = appointment.Organizer.EmailAddress,
                Activities = new List<Activity>
                {
                    new () { Name = appointment.Subject, CapacityPerDay = 8 } // Assuming each appointment takes 8 hours
                }
            };

            capacityItems.Add(capacityItem);
        }

        const string requestUrl = $"{AzureDevOpsUri}/{TeamId}/_apis/work/teamsettings/iterations/{IterationId}/capacities?api-version=6.0";

        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
        var requestBody = JsonSerializer.Serialize(capacityItems, options);
        var content = new StringContent(requestBody, Encoding.UTF8, "application/json");

        var response = await httpClient.PostAsync(requestUrl, content);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new Exception($"Failed to post appointments to Azure DevOps sprint capacity. Error: {responseContent}");
        }
    }
}

internal class GraphCalendarEventsResponse
{
    public List<Appointment> Value { get; set; }
}

internal class Appointment
{
    public string Subject { get; set; }
    public GraphEmailAddress Organizer { get; set; }
}

internal class GraphEmailAddress
{
    public string EmailAddress { get; set; }
}

internal class CapacityItem
{
    public string TeamId { get; set; }
    public string IterationId { get; set; }
    public string TeamMemberEmail { get; set; }
    public List<Activity> Activities { get; set; }
}

internal class Activity
{
    public string Name { get; set; }
    public int CapacityPerDay { get; set; }
}

internal class AzureDevOpsIteration
{
    public AzureDevOpsIterationAttributes Attributes { get; set; }
}

internal class AzureDevOpsIterationAttributes
{
    public DateTime StartDate { get; set; }
    public DateTime FinishDate { get; set; }
}