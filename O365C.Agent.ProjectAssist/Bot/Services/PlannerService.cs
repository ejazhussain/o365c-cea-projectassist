using Azure.Core;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using O365C.Agent.ProjectAssist.Bot.Helpers;
using O365C.Agent.ProjectAssist.Bot.Models;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace O365C.Agent.ProjectAssist.Bot.Services;

public interface IPlannerService
{
    Task<List<PlannerTaskResponse>> GetTasksAsync(string accessToken);
    Task<List<PlannerTaskResponse>> GetTasksForUserAsync(string accessToken, string email);
    Task<List<PlannerTaskResponse>> GetTasksByPriorityAsync(string accessToken, string email, int priority);
    Task<List<PlannerTaskResponse>> GetTasksByProgressAsync(string accessToken, string email, string progressStatus);
    Task<List<PlannerTaskResponse>> GetOverdueTasksByEmailAsync(string accessToken, string email);
    Task<List<PlannerTaskResponse>> GetOverdueTasksAsync(string accessToken);
    Task<User> GetUserByEmailAsync(string accessToken, string email);
    Task<PlannerPlan> GetPlanAsync(string accessToken, string planName);
    Task<PlannerBucket> GetBucketAsync(string accessToken, string planId);
    Task<PlannerTaskResponse> CreateTaskAsync(string accessToken, string planId, string bucketId, string title, string email = null, string description = null, DateTimeOffset? dueDateTime = null);
    Task<List<PlannerTaskResponse>> GetTasksForUserInPlanAsync(string accessToken, string email, string planId);
}

public class PlannerService : IPlannerService
{

    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ConfigOptions _configOptions;
    // The constructor now takes an IAuthenticationProvider, which is more flexible.
    // However, for user-specific tokens like "On-Behalf-Of", you'll typically pass a TokenCredential.
    public PlannerService(IHttpClientFactory httpClientFactory, ConfigOptions configOptions)
    {
        _httpClientFactory = httpClientFactory ?? throw new ArgumentNullException(nameof(httpClientFactory));
        _configOptions = configOptions ?? throw new ArgumentNullException(nameof(configOptions));
    }

    /// <summary>
    /// Gets all planner tasks for the current user using Microsoft Graph.
    /// </summary>
    public async Task<List<PlannerTaskResponse>> GetTasksAsync(string accessToken)
    {
        if (string.IsNullOrWhiteSpace(accessToken))
            throw new ArgumentException("Access token cannot be null or empty.", nameof(accessToken));

        try
        {
            var graphClient = GraphAuthHelper.CreateGraphClientWithAccessToken(accessToken);
            var plannerTasks = await graphClient.Me.Planner.Tasks
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Expand = new[] { "details" };
                }).ConfigureAwait(false);

            var result = new List<PlannerTaskResponse>();

            if (plannerTasks?.Value != null)
            {
                foreach (var task in plannerTasks.Value)
                {
                    var assignments = new Dictionary<string, PlannerAssignment>();
                    if (task.Assignments != null)
                    {
                        foreach (var assignment in task.Assignments.AdditionalData)
                        {
                            if (assignment.Value is PlannerAssignment plannerAssignment)
                            {
                                assignments.Add(assignment.Key, plannerAssignment);
                            }
                        }
                    }
                    result.Add(new PlannerTaskResponse
                    {
                        Id = task.Id,
                        PlanId = task.PlanId,
                        BucketId = task.BucketId,
                        Title = task.Title,
                        OrderHint = task.OrderHint,
                        AssigneePriority = task.AssigneePriority,
                        CreatedDateTime = task.CreatedDateTime ?? default,
                        CreatedBy = task.CreatedBy,
                        HasDescription = task.HasDescription ?? false,
                        Details = task.Details != null ? new PlannerTaskDetails
                        {
                            Id = task.Details.Id,
                            Description = task.Details.Description,
                            PreviewType = task.Details.PreviewType,
                            References = task.Details.References,
                            Checklist = task.Details.Checklist
                        } : null,
                        PercentComplete = task.PercentComplete ?? 0,
                        Priority = task.Priority ?? 0,
                        StartDateTime = task.StartDateTime,
                        DueDateTime = task.DueDateTime,
                        Assignments = assignments
                    });
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            // TODO: Replace with ILogger for production
            Console.Error.WriteLine($"Error in GetTasksAsync: {ex.Message}");
            throw new Exception("Failed to retrieve planner tasks.", ex);
        }
    }

    /// <summary>
    /// Gets all planner tasks assigned to a user by email using Microsoft Graph.
    /// </summary>
    public async Task<List<PlannerTaskResponse>> GetTasksForUserAsync(string accessToken, string email)
    {
        if (string.IsNullOrWhiteSpace(accessToken))
            throw new ArgumentException("Access token cannot be null or empty.", nameof(accessToken));
        if (string.IsNullOrWhiteSpace(email))
            throw new ArgumentException("Email cannot be null or empty.", nameof(email));

        try
        {
            var user = await GetUserByEmailAsync(accessToken, email).ConfigureAwait(false);
            if (user == null || string.IsNullOrEmpty(user.Id))
            {
                return new List<PlannerTaskResponse>();
            }

            var graphClient = GraphAuthHelper.CreateGraphClientWithAccessToken(accessToken);
            var plannerTasks = await graphClient.Users[email].Planner.Tasks
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Expand = new[] { "details" };
                }).ConfigureAwait(false);

            var result = new List<PlannerTaskResponse>();

            if (plannerTasks?.Value != null)
            {
                foreach (var task in plannerTasks.Value)
                {
                    // Check if the task is assigned to the specified user
                    bool isAssignedToUser = false;
                    if (task.Assignments != null && task.Assignments.AdditionalData != null)
                    {
                        isAssignedToUser = task.Assignments.AdditionalData.ContainsKey(user.Id);
                    }
                    if (isAssignedToUser)
                    {
                        var assignments = new Dictionary<string, PlannerAssignment>();
                        if (task.Assignments != null)
                        {
                            foreach (var assignment in task.Assignments.AdditionalData)
                            {
                                if (assignment.Value is PlannerAssignment plannerAssignment)
                                {
                                    assignments.Add(assignment.Key, plannerAssignment);
                                }
                            }
                        }
                        result.Add(new PlannerTaskResponse
                        {
                            Id = task.Id,
                            PlanId = task.PlanId,
                            BucketId = task.BucketId,
                            Title = task.Title,
                            OrderHint = task.OrderHint,
                            AssigneePriority = task.AssigneePriority,
                            CreatedDateTime = task.CreatedDateTime ?? default,
                            CreatedBy = task.CreatedBy,
                            HasDescription = task.HasDescription ?? false,
                            Details = task.Details != null ? new PlannerTaskDetails
                            {
                                Id = task.Details.Id,
                                Description = task.Details.Description,
                                PreviewType = task.Details.PreviewType,
                                References = task.Details.References,
                                Checklist = task.Details.Checklist
                            } : null,
                            PercentComplete = task.PercentComplete ?? 0,
                            Priority = task.Priority ?? 0,
                            StartDateTime = task.StartDateTime,
                            DueDateTime = task.DueDateTime,
                            Assignments = assignments
                        });
                    }
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            // TODO: Replace with ILogger for production
            Console.Error.WriteLine($"Error in GetTasksForUserAsync: {ex.Message}");
            throw new Exception("Failed to retrieve planner tasks for user.", ex);
        }
    }

    /// <summary>
    /// Gets all planner tasks assigned to a user by email in a specific plan using Microsoft Graph.
    /// </summary>
    public async Task<List<PlannerTaskResponse>> GetTasksForUserInPlanAsync(string accessToken, string email, string planId)
    {
        try
        {

            var userTasks = await GetTasksForUserAsync(accessToken, email);
            if (string.IsNullOrWhiteSpace(planId))
            {
                return userTasks;
            }
            var result = userTasks.Where(t => t.PlanId.Equals(planId, StringComparison.OrdinalIgnoreCase)).ToList();
            return result;

        }
        catch (Exception ex)
        {
            // TODO: Replace with ILogger for production
            Console.Error.WriteLine($"Error in GetTasksForUserInPlanAsync: {ex.Message}");
            throw new Exception("Failed to retrieve planner tasks for user in plan.", ex);
        }
    }

    // Fetches all Planner tasks assigned to a specific user with a given priority
    public async Task<List<PlannerTaskResponse>> GetTasksByPriorityAsync(string accessToken, string email, int priority)
    {
        var userTasks = await GetTasksForUserAsync(accessToken, email);
        return userTasks.Where(t => t.Priority == priority).ToList();
    }

    // Fetches all Planner tasks assigned to a specific user that are not completed
    public async Task<List<PlannerTaskResponse>> GetTasksByProgressAsync(string accessToken, string email, string progressStatus)
    {
        var userTasks = await GetTasksForUserAsync(accessToken, email);
        switch (progressStatus?.Trim().ToLowerInvariant())
        {
            case "notstarted":
            case "not started":
                return userTasks.Where(t => t.PercentComplete == 0).ToList();
            case "inprogress":
            case "in progress":
            case "in-progress":
                return userTasks.Where(t => t.PercentComplete > 0 && t.PercentComplete < 100).ToList();
            case "completed":
            case "done":
                return userTasks.Where(t => t.PercentComplete >= 100).ToList();
            case "incomplete":
            case "open":
                return userTasks.Where(t => t.PercentComplete < 100).ToList();
            default:
                return userTasks;
        }
    }

    // Fetches all Planner tasks assigned to a specific user that are overdue (due date expired and not completed)
    public async Task<List<PlannerTaskResponse>> GetOverdueTasksByEmailAsync(string accessToken, string email)
    {
        var userTasks = await GetTasksForUserAsync(accessToken, email);
        var now = DateTimeOffset.UtcNow;
        return userTasks.Where(t => t.DueDateTime.HasValue && t.DueDateTime.Value < now && t.PercentComplete < 100).ToList();
    }
    //Fetch all overdue tasks
    public async Task<List<PlannerTaskResponse>> GetOverdueTasksAsync(string accessToken)
    {
        var allTasks = await GetTasksAsync(accessToken);
        var now = DateTimeOffset.UtcNow;
        return allTasks.Where(t => t.DueDateTime.HasValue && t.DueDateTime.Value < now && t.PercentComplete < 100).ToList();
    }

    // Fetches a user object by email address using Microsoft Graph
    public async Task<User> GetUserByEmailAsync(string accessToken, string email)
    {
        if (string.IsNullOrEmpty(accessToken))
        {
            throw new ArgumentException("Access token cannot be null or empty.", nameof(accessToken));
        }
        if (string.IsNullOrEmpty(email))
        {
            throw new ArgumentException("Email cannot be null or empty.", nameof(email));
        }

        GraphServiceClient graphClient = GraphAuthHelper.CreateGraphClientWithClientSecret(
            _configOptions.MicrosoftGraph.TenantId,
            _configOptions.MicrosoftGraph.ClientId,
            _configOptions.MicrosoftGraph.ClientSecret,
            new[] { "https://graph.microsoft.com/.default" });
        var users = await graphClient.Users
            .GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Filter = $"mail eq '{email}' or userPrincipalName eq '{email}'";
                requestConfiguration.QueryParameters.Top = 1;
            });

        return users?.Value?.FirstOrDefault();
    }

    //Add new method to create new task under specified plan and bucket
    public async Task<PlannerTaskResponse> CreateTaskAsync(string accessToken, string planId, string bucketId, string title, string email, string description = null, DateTimeOffset? dueDateTime = null)
    {
        if (string.IsNullOrWhiteSpace(accessToken))
            throw new ArgumentException("Access token cannot be null or empty.", nameof(accessToken));
        if (string.IsNullOrWhiteSpace(planId))
            throw new ArgumentException("Plan ID cannot be null or empty.", nameof(planId));
        if (string.IsNullOrWhiteSpace(bucketId))
            throw new ArgumentException("Bucket ID cannot be null or empty.", nameof(bucketId));
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Title cannot be null or empty.", nameof(title));
        // if (string.IsNullOrWhiteSpace(email))
        //     throw new ArgumentException("Email cannot be null or empty.", nameof(email));

        //get user by email to get user ID for assignment
        // var user = await GetUserByEmailAsync(accessToken, email).ConfigureAwait(false);
        // if (user == null || string.IsNullOrEmpty(user.Id))
        // {
        //     throw new Exception($"User with email {email} not found.");
        // }

        var graphClient = GraphAuthHelper.CreateGraphClientWithAccessToken(accessToken);
        var requestBody = new PlannerTask
        {
            PlanId = planId,
            BucketId = bucketId,
            Title = title,
            // Assignments = new PlannerAssignments
            // {
            //     AdditionalData = new Dictionary<string, object>
            //     {
            //         {
            //             user.Id , new PlannerAssignment
            //             {
            //                 OdataType = "#microsoft.graph.plannerAssignment",
            //                 OrderHint = " !",
            //             }
            //         },
            //     },
            // },
            //Details = new PlannerTaskDetails
            //{
            //    Description = description
            //},
            //DueDateTime = dueDateTime
        };

        try
        {

            var createdTask = await graphClient.Planner.Tasks.PostAsync(requestBody);

            return new PlannerTaskResponse
            {
                Id = createdTask.Id,
                PlanId = createdTask.PlanId,
                BucketId = createdTask.BucketId,
                Title = createdTask.Title,
                Details = createdTask.Details != null ? new PlannerTaskDetails
                {
                    Id = createdTask.Details.Id,
                    Description = createdTask.Details.Description,
                    PreviewType = createdTask.Details.PreviewType,
                    References = createdTask.Details.References,
                    Checklist = createdTask.Details.Checklist
                } : null,
                DueDateTime = createdTask.DueDateTime
            };
        }
        catch (Exception ex)
        {
            // TODO: Replace with ILogger for production
            Console.Error.WriteLine($"Error in CreateTaskAsync: {ex.Message}");
            throw new Exception("Failed to create planner tasks", ex);

        }
    }

    //Get all plans for the current user
    public async Task<PlannerPlan> GetPlanAsync(string accessToken, string planName)
    {
        if (string.IsNullOrWhiteSpace(accessToken))
            throw new ArgumentException("Access token cannot be null or empty.", nameof(accessToken));
        try
        {
            var graphClient = GraphAuthHelper.CreateGraphClientWithAccessToken(accessToken);
            var plans = await graphClient.Me.Planner.Plans.GetAsync().ConfigureAwait(false);
            var planstList = plans?.Value?.ToList() ?? new List<PlannerPlan>();
            //filter plans by name
            PlannerPlan result = null;
            if (!string.IsNullOrWhiteSpace(planName))
            {
                result = planstList.FirstOrDefault(p => p.Title.Contains(planName, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                result = planstList.FirstOrDefault();
            }
            return result ?? throw new Exception($"No plans found with name {planName}.");


        }
        catch (Exception ex)
        {
            //TODO: Replace with ILogger for production
            Console.Error.WriteLine($"Error in GetPlansAsync: {ex.Message}");
            throw new Exception("Failed to retrieve planner plans.", ex);
        }
    }
    //Get all buckets for a specific plan
    public async Task<PlannerBucket> GetBucketAsync(string accessToken, string planId)
    {
        if (string.IsNullOrWhiteSpace(accessToken))
            throw new ArgumentException("Access token cannot be null or empty.", nameof(accessToken));
        if (string.IsNullOrWhiteSpace(planId))
            throw new ArgumentException("Plan ID cannot be null or empty.", nameof(planId));

        try
        {
            var graphClient = GraphAuthHelper.CreateGraphClientWithAccessToken(accessToken);
            var buckets = await graphClient.Planner.Plans[planId].Buckets.GetAsync().ConfigureAwait(false);
            var bucketList = buckets?.Value?.ToList() ?? new List<PlannerBucket>();
            //filter buckets by planId
            var result = bucketList.FirstOrDefault(b => b.PlanId.Equals(planId, StringComparison.OrdinalIgnoreCase));
            return result ?? throw new Exception($"No buckets found for plan ID {planId}.");

        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error in GetBucketsAsync: {ex.Message}");
            throw new Exception("Failed to retrieve planner buckets.", ex);

        }
    }
}
