using Microsoft.Agents.Authentication.Msal;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.App.UserAuth;
using Microsoft.Agents.Builder.UserAuth;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Storage;
using Microsoft.Extensions.Logging;
using Microsoft.SemanticKernel;
using O365C.Agent.ProjectAssist;
using O365C.Agent.ProjectAssist.Bot.Agents;
using O365C.Agent.ProjectAssist.Bot.Services;


var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();
builder.Logging.AddConsole();


// Register Semantic Kernel
builder.Services.AddKernel();

// Register the AI service of your choice. AzureOpenAI and OpenAI are demonstrated...
var config = builder.Configuration.Get<ConfigOptions>();

builder.Services.AddAzureOpenAIChatCompletion(
    deploymentName: config.Azure.OpenAIDeploymentName,
    endpoint: config.Azure.OpenAIEndpoint,
    apiKey: config.Azure.OpenAIApiKey
);

// Add our configuration class                
builder.Services.AddSingleton(options => { return config; });

// Register the WeatherForecastAgent
//builder.Services.AddTransient<WeatherForecastAgent>();

builder.Services.AddTransient<ProjectAssistAgent>();


// Register PlannerService and IPlannerService
builder.Services.AddTransient<IPlannerService, PlannerService>();
builder.Services.AddTransient<IEmailService, EmailService>();


//builder.Services.AddHttpClient<IPlannerService, PlannerService>(client =>
//{
//    // Configure the HttpClient for PlannerService if needed
//    client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
//    client.Timeout = TimeSpan.FromSeconds(60);
//});

// Add AspNet token validation
builder.Services.AddBotAspNetAuthentication(builder.Configuration);

// Register IStorage.  For development, MemoryStorage is suitable.
// For production Agents, persisted storage should be used so
// that state survives Agent restarts, and operate correctly
// in a cluster of Agent instances.
builder.Services.AddSingleton<IStorage, MemoryStorage>();


// Add AgentApplicationOptions from config.
builder.AddAgentApplicationOptions();

// Add AgentApplicationOptions.  This will use DI'd services and IConfiguration for construction.
builder.Services.AddTransient<AgentApplicationOptions>();

// Add the bot (which is transient)
// builder.AddAgent<O365C.Agent.ProjectAssist.Bot.WeatherAgentBot>();
builder.AddAgent<O365C.Agent.ProjectAssist.Bot.ProjectAssistBot>();

// ====================================================================
// CONFIGURE AgentApplicationOptions and UserAuthorization Handlers
// ==================================================== ================

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapPost("/api/messages", async (HttpRequest request, HttpResponse response, IAgentHttpAdapter adapter, IAgent agent, CancellationToken cancellationToken) =>
{
    await adapter.ProcessAsync(request, response, agent, cancellationToken);
});

if (app.Environment.IsDevelopment() || app.Environment.EnvironmentName == "Playground")
{
    app.MapGet("/", () => "Weather Bot");
    app.UseDeveloperExceptionPage();
    app.MapControllers().AllowAnonymous();
}
else
{
    app.MapControllers();
}

app.Run();

