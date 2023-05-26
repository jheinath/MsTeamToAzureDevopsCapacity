using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Identity.Client;

class Program
{
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

    static async Task Main(string[] args)
    {
        // Authenticate and get access token for Microsoft Graph API
        var graphAccessToken = await GetAccessToken();

        // Get start and end dates of the sprint
        DateTime sprintStartDate = await GetSprintStartDate();
        DateTime sprintEndDate = await GetSprintEndDate();

        // Retrieve appointments from the specified Microsoft Teams channel calendar within the sprint range
        var appointments = await GetAppointments(graphAccessToken, sprintStartDate, sprintEndDate);

        // Create absences for each team member based on their appointments
        var absences = CreateAbsences(appointments);

        // Copy capacity from the previous sprint
        var previousCapacity = await GetSprintCapacity(TeamId, await GetPreviousSprintId());

        // Update the capacity with the absences
        var updatedCapacity = UpdateCapacityWithAbsences(previousCapacity, absences);

        // Post updated capacity to Azure DevOps sprint capacity for each team member
        await PostCapacityToSprint(updatedCapacity);

        Console.WriteLine("Capacity posted to Azure DevOps sprint successfully.");
    }

    static async Task<string> GetAccessToken()
    {
        var clientApp = ConfidentialClientApplicationBuilder.Create(ClientId)
            .WithClientSecret(ClientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{TenantId}/v2.0"))
            .Build();

        var authResult = await clientApp.AcquireTokenForClient(new string[] { $"{GraphApiUrl}/.default" })
            .ExecuteAsync();

        return authResult.AccessToken;
    }

    static async Task<DateTime> GetSprintStartDate()
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", PatToken);

        var requestUrl = $"{AzureDevOpsUri}/{TeamId}/_apis/work/teamsettings/iterations/{IterationId}?api-version=6.0";
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

    static async Task<DateTime> GetSprintEndDate()
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", PatToken);

        var requestUrl = $"{AzureDevOpsUri}/{TeamId}/_apis/work/teamsettings/iterations/{IterationId}?api-version=6.0";
        var response = await httpClient.GetAsync(requestUrl);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (response.IsSuccessStatusCode)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var iteration = JsonSerializer.Deserialize<AzureDevOpsIteration>(responseContent, options);
            return iteration.Attributes.FinishDate;
        }
        else
        {
            throw new Exception($"Failed to retrieve sprint end date from Azure DevOps. Error: {responseContent}");
        }
    }

    static async Task<List<Appointment>> GetAppointments(string accessToken, DateTime startDate, DateTime endDate)
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        string startDateTime = startDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
        string endDateTime = endDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
        string requestUrl = $"{GraphApiUrl}/groups/{TeamsGroupId}/channels/{ChannelId}/calendarView?startDateTime={startDateTime}&endDateTime={endDateTime}";

        var response = await httpClient.GetAsync(requestUrl);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (response.IsSuccessStatusCode)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var events = JsonSerializer.Deserialize<GraphCalendarEventsResponse>(responseContent, options);
            return events.Value;
        }
        else
        {
            throw new Exception($"Failed to retrieve appointments from Microsoft Teams channel calendar. Error: {responseContent}");
        }
    }

    static List<Absence> CreateAbsences(List<Appointment> appointments)
    {
        var absences = new List<Absence>();

        foreach (var appointment in appointments)
        {
            var absence = new Absence
            {
                TeamId = TeamId,
                IterationId = IterationId,
                TeamMemberEmail = appointment.Organizer.Email,
                StartDate = appointment.Start.DateTime,
                EndDate = appointment.End.DateTime
            };

            absences.Add(absence);
        }

        return absences;
    }

    static async Task<List<SprintCapacity>> GetSprintCapacity(string teamId, string sprintId)
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", PatToken);

        var requestUrl = $"{AzureDevOpsUri}/{teamId}/_apis/work/teamsettings/iterations/{sprintId}/capacities?api-version=6.0";

        var response = await httpClient.GetAsync(requestUrl);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (response.IsSuccessStatusCode)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var capacityResponse = JsonSerializer.Deserialize<AzureDevOpsCapacityResponse>(responseContent, options);
            return capacityResponse.Capacities;
        }
        else
        {
            throw new Exception($"Failed to retrieve sprint capacity from Azure DevOps. Error: {responseContent}");
        }
    }

    static List<SprintCapacity> UpdateCapacityWithAbsences(List<SprintCapacity> capacity, List<Absence> absences)
    {
        var updatedCapacity = new List<SprintCapacity>();

        foreach (var cap in capacity)
        {
            var absence = absences.FirstOrDefault(a => a.TeamMemberEmail.Equals(cap.TeamMemberEmail));

            if (absence != null)
            {
                cap.Activity = "Absence";
                cap.CapacityPerDay = 0;
                cap.StartDate = absence.StartDate.Date;
                cap.EndDate = absence.EndDate.Date.AddDays(1).AddTicks(-1);
            }

            updatedCapacity.Add(cap);
        }

        return updatedCapacity;
    }

    static async Task PostCapacityToSprint(List<SprintCapacity> capacity)
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", PatToken);

        var requestUrl = $"{AzureDevOpsUri}/{TeamId}/_apis/work/teamsettings/iterations/{IterationId}/capacities?api-version=6.0";

        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };
        var requestBody = JsonSerializer.Serialize(capacity, options);
        var content = new StringContent(requestBody, Encoding.UTF8, "application/json");

        var response = await httpClient.PostAsync(requestUrl, content);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new Exception($"Failed to post capacity to Azure DevOps sprint capacity. Error: {responseContent}");
        }
    }

    static async Task<string> GetPreviousSprintId()
    {
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", PatToken);

        var requestUrl = $"{AzureDevOpsUri}/{TeamId}/_apis/work/teamsettings/iterations/{IterationId}?api-version=6.0-preview.1";
        var response = await httpClient.GetAsync(requestUrl);
        var responseContent = await response.Content.ReadAsStringAsync();

        if (response.IsSuccessStatusCode)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var iteration = JsonSerializer.Deserialize<AzureDevOpsIteration>(responseContent, options);
            return iteration.Relations.Single().Id;
        }
        else
        {
            throw new Exception($"Failed to retrieve previous sprint ID from Azure DevOps. Error: {responseContent}");
        }
    }
}

class GraphCalendarEventsResponse
{
    public List<Appointment> Value { get; set; }
}

class Appointment
{
    public string Subject { get; set; }
    public GraphEmailAddress Organizer { get; set; }
    public AppointmentDateTime Start { get; set; }
    public AppointmentDateTime End { get; set; }
}

class GraphEmailAddress
{
    public string Email { get; set; }
}

class AppointmentDateTime
{
    public DateTime DateTime { get; set; }
}

class Absence
{
    public string TeamId { get; set; }
    public string IterationId { get; set; }
    public string TeamMemberEmail { get; set; }
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
}

class AzureDevOpsCapacityResponse
{
    public List<SprintCapacity> Capacities { get; set; }
}

class SprintCapacity
{
    public string TeamMemberEmail { get; set; }
    public string Activity { get; set; }
    public double CapacityPerDay { get; set; }
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
}

class AzureDevOpsIteration
{
    public List<AzureDevOpsRelation> Relations { get; set; }
    public AzureDevOpsAttributes Attributes { get; set; }
}

class AzureDevOpsAttributes
{
    public DateTime StartDate { get; set; }
    public DateTime FinishDate { get; set; }
}

class AzureDevOpsRelation
{
    public string Id { get; set; }
    public string Rel { get; set; }
    public string Url { get; set; }
}
