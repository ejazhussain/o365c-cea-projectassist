using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App;
using Microsoft.SemanticKernel;
using O365C.Agent.ProjectAssist.Bot.Services;
using System.ComponentModel;
using System.Threading.Tasks;

namespace O365C.Agent.ProjectAssist.Bot.Plugins;

public class TeamsOutlookPlugin
{
    private readonly IEmailService _emailService;
    private readonly AgentApplication _app;
    private readonly ITurnContext _turnContext;
    public TeamsOutlookPlugin(AgentApplication app, ITurnContext turnContext, IEmailService emailService)
    {
        _app = app;
        _turnContext = turnContext;
        _emailService = emailService;
    }

    /// <summary>
    /// Sends an email using the configured email service.
    /// /// </summary>    
    /// <param name="toEmail">The email address of the recipient.</param>
    /// <param name="subject">The subject of the email.</param>
    /// <param name="body">The body content of the email.</param>
    /// 
    [KernelFunction, Description("Send an email using the configured email service.")]
    public async Task<bool> SendEmailAsync(string toEmail, string subject, string body)
    {
        var accessToken = await _app.UserAuthorization.GetTurnTokenAsync(_turnContext, handlerName: "graph");
        var fromEmail = "ehussain@office365clinic.com";
        return await _emailService.SendEmailAsync(accessToken, fromEmail, toEmail, subject, body);
    }    
}
