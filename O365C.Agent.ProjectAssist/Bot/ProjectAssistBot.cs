using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.Agents.Builder.State;
using Microsoft.Agents.Core.Models;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using O365C.Agent.ProjectAssist.Bot.Agents;
using O365C.Agent.ProjectAssist.Bot.Models;
using O365C.Agent.ProjectAssist.Bot.Services;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;

namespace O365C.Agent.ProjectAssist.Bot;

public class ProjectAssistBot : AgentApplication
{
    private ProjectAssistAgent _projectAssistAgent;
    private Kernel _kernel;
    private readonly IHttpClientFactory _httpClientFactory; // Add this field
    private readonly IPlannerService _plannerService;
    private readonly IEmailService _emailService;
    public ConfigOptions _configurations; 

    public ProjectAssistBot(AgentApplicationOptions options, Kernel kernel, IHttpClientFactory httpClientFactory, IPlannerService graphService, ConfigOptions configurations, IEmailService emailService) : base(options)
    {
        _kernel = kernel ?? throw new ArgumentNullException(nameof(kernel));
        _httpClientFactory = httpClientFactory ?? throw new ArgumentNullException(nameof(httpClientFactory));
        OnConversationUpdate(ConversationUpdateEvents.MembersAdded, WelcomeMessageAsync);
        OnActivity(ActivityTypes.Message, MessageActivityAsync, rank: RouteRank.Last);
        _plannerService = graphService;
        _configurations = configurations;
        _emailService = emailService;
    }

    protected async Task MessageActivityAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {


        // Setup local service connection
        ServiceCollection serviceCollection = [
            new ServiceDescriptor(typeof(ITurnState), turnState),
            new ServiceDescriptor(typeof(ITurnContext), turnContext),
            new ServiceDescriptor(typeof(Kernel), _kernel),
            new ServiceDescriptor(typeof(AgentApplication), this),
            new ServiceDescriptor(typeof(IHttpClientFactory), _httpClientFactory),
            new ServiceDescriptor(typeof(IPlannerService), _plannerService),
            new ServiceDescriptor(typeof(IEmailService), _emailService),
            new ServiceDescriptor(typeof(ConfigOptions), _configurations),
        ];

        await turnContext.StreamingResponse.QueueInformativeUpdateAsync("Working on a response for you");

        ChatHistory chatHistory = turnState.GetValue("conversation.chatHistory", () => new ChatHistory());
        _projectAssistAgent = new ProjectAssistAgent(_kernel, serviceCollection.BuildServiceProvider());

        var response = await _projectAssistAgent.InvokeAgentAsync(turnContext.Activity.Text, chatHistory);
        if (response == null)
        {
            turnContext.StreamingResponse.QueueTextChunk("Sorry, I couldn't get a project management response at the moment.");
            await turnContext.StreamingResponse.EndStreamAsync(cancellationToken);
            return;
        }

        switch (response.ContentType)
        {
            case ProjectAssistAgentResponseContentType.Text:
                turnContext.StreamingResponse.QueueTextChunk(response.Content);
                break;
            case ProjectAssistAgentResponseContentType.AdaptiveCard:
                turnContext.StreamingResponse.FinalMessage = MessageFactory.Attachment(new Attachment()
                {
                    ContentType = "application/vnd.microsoft.card.adaptive",
                    Content = response.Content,
                });
                break;
            default:
                break;
        }
        await turnContext.StreamingResponse.EndStreamAsync(cancellationToken);
    }

    protected async Task WelcomeMessageAsync(ITurnContext turnContext, ITurnState turnState, CancellationToken cancellationToken)
    {
        foreach (ChannelAccount member in turnContext.Activity.MembersAdded)
        {
            if (member.Id != turnContext.Activity.Recipient.Id)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("Hello and Welcome! I'm here to help with your project management needs!"), cancellationToken);
            }
        }
    }


   

}
