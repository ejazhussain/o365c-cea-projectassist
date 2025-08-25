using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.UserAuth;
using Microsoft.Graph.Models;
using Microsoft.SemanticKernel;
using O365C.Agent.ProjectAssist.Bot.Models;
using O365C.Agent.ProjectAssist.Bot.Services;
using System.ComponentModel;
using System.Text.Json.Serialization;

namespace O365C.Agent.ProjectAssist.Bot.Plugins;

public class PlannerPlugin
{
    private readonly IPlannerService _plannerService;
    private readonly AgentApplication _app;
    private readonly ITurnContext _turnContext;
    public PlannerPlugin(AgentApplication app, ITurnContext turnContext, IPlannerService plannerService)
    {
        _app = app;
        _turnContext = turnContext;
        _plannerService = plannerService;
    }


    /// <summary>
    /// Gets all planner tasks for the current user using Microsoft Graph.
    /// </summary>
    [KernelFunction, Description("Get all planner tasks for the current user.")]
    public async Task<List<PlannerTaskResponse>> GetAllPlannerTasksAsync()
    {
        try
        {
            var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
            var tasks = await _plannerService.GetTasksAsync(accessToken);
            return tasks;
        }
        catch (Exception ex)
        {
            throw new Exception("Failed to retrieve planner tasks.", ex);
        }
    }

    /// <summary>
    /// Gets all planner tasks assigned to a user by email.
    /// </summary>
    [KernelFunction, Description("Get all planner tasks assigned to a user by email.")]
    public async Task<List<PlannerTaskResponse>> GetTasksForUserAsync(string email)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        return await _plannerService.GetTasksForUserAsync(accessToken, email);
    }

    /// <summary>
    /// Gets all planner tasks assigned to a user by email and priority.
    /// </summary>
    [KernelFunction, Description("Get all planner tasks assigned to a user by email and priority.")]
    public async Task<List<PlannerTaskResponse>> GetTasksByPriorityAsync(string email, int priority)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        return await _plannerService.GetTasksByPriorityAsync(accessToken, email, priority);
    }

    /// <summary>
    /// Gets all planner tasks assigned to a user by email and progress status (e.g., NotStarted, InProgress, Completed).
    /// </summary>
    [KernelFunction, Description("Get all planner tasks assigned to a user by email and progress status (NotStarted, InProgress, Completed, etc.).")]
    public async Task<List<PlannerTaskResponse>> GetTasksByProgressAsync(string email, string progressStatus)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        return await _plannerService.GetTasksByProgressAsync(accessToken, email, progressStatus);
    }

    /// <summary>
    /// Gets all overdue planner tasks assigned to a user by email.
    /// </summary>
    [KernelFunction, Description("Get all overdue planner tasks assigned to a user by email.")]
    public async Task<List<PlannerTaskResponse>> GetOverdueTasksAsync(string email)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        return await _plannerService.GetOverdueTasksByEmailAsync(accessToken, email);
    }

    /// <summary>
    /// Gets all overdue planner tasks across all users (no email filter).
    /// </summary>
    [KernelFunction, Description("Get all overdue planner tasks across all users.")]
    public async Task<List<PlannerTaskResponse>> GetAllOverdueTasksAsync()
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        return await _plannerService.GetOverdueTasksAsync(accessToken);
    }

    /// <summary>
    /// Gets a planner plan for the current user by plan name.
    /// </summary>
    [KernelFunction, Description("Get a planner plan for the current user by plan name.")]
    public async Task<PlannerPlan> GetPlanAsync(string planName)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        return await _plannerService.GetPlanAsync(accessToken, planName);
    }

    /// <summary>
    /// Gets a planner bucket for a specific plan by plan ID.
    /// </summary>
    [KernelFunction, Description("Get a planner bucket for a specific plan by plan ID. Use GetPlanAsync method to fetch the planId")]
    public async Task<PlannerBucket> GetBucketAsync(string planId)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        return await _plannerService.GetBucketAsync(accessToken, planId);
    }

    /// <summary>
    /// Creates a new planner task in a specific plan and bucket.
    /// </summary>
    /// <param name="planName">The name of the plan where the task will be created.</param>    
    /// <param name="title">The title of the task.</param>
    /// <param name="email">The email of the user to assign the task to (optional).</param>
    /// <param name="description">The description of the task (optional).</param>   
    /// <param name="dueDateTime">The due date and time for the task (optional).</param>
    [KernelFunction, Description("Create a new planner task in a specific plan and bucket. use GetPlanAsync method to fetch detail about the plan, and use GetBucketsAsync to fetch detail about the bucket assoicated with the given plan ")]
    public async Task<PlannerTaskResponse> CreateTaskAsync(string planName, string title, string email = null, string description = null, DateTimeOffset? dueDateTime = null)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        var plan = await _plannerService.GetPlanAsync(accessToken, planName);
        if (plan == null)
            throw new Exception($"Plan '{planName}' not found.");

        var bucket = await _plannerService.GetBucketAsync(accessToken, plan.Id);
        if (bucket == null)
            throw new Exception($"Bucket not found in plan '{planName}'.");

        return await _plannerService.CreateTaskAsync(accessToken, plan.Id, bucket.Id, title, email, description, dueDateTime);
    }

    /// <summary>
    /// Gets all planner tasks assigned to a user by email in a specific plan.
    /// </summary>
    /// <param name="email">The email of the user to get tasks for.</param>
    /// <param name="planName">The Name of the plan to filter tasks by.</param>
    /// 
    [KernelFunction, Description("Get all planner tasks assigned to a user by email in a specific plan.")]
    public async Task<List<PlannerTaskResponse>> GetTasksForUserInPlanAsync(string email, string planName)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
         var plan = await _plannerService.GetPlanAsync(accessToken, planName);
        if (plan == null)
        {
            throw new Exception($"Plan '{planName}' not found.");
        }
        return await _plannerService.GetTasksForUserInPlanAsync(accessToken, email, plan.Id);
    }

}
