using Microsoft.SemanticKernel.Connectors.OpenAI;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.Agents;
using Microsoft.SemanticKernel.ChatCompletion;
using System.Text;
using System.Text.Json.Nodes;
using O365C.Agent.ProjectAssist.Bot.Plugins;
using O365C.Agent.ProjectAssist.Bot.Models;

namespace O365C.Agent.ProjectAssist.Bot.Agents
{
    public class ProjectAssistAgent
    {
        private readonly Kernel _kernel;
        private readonly ChatCompletionAgent _agent;

        private const string AgentName = "ProjectAssistAgent";
        private const string AgentInstructions = """
            You are a helpful assistant for project management tasks. You help users manage tasks, projects, and collaborate using Microsoft 365 tools like Planner, SharePoint, Teams, and Outlook.
            You can also send emails, schedule meetings, and manage calendar events using Teams and Outlook via the TeamsOutlookPlugin.
            Ask follow-up questions to clarify requirements. When you have enough information, respond with a summary or actionable steps, formatted as an adaptive card if appropriate.
            Use adaptive cards version 1.5 or later for visual responses.
            Respond in JSON format with the following schema:
            {
                \"contentType\": "'Text' or 'AdaptiveCard' only",
                \"content\": "{The content of the response, may be plain text, or JSON based adaptive card}"
            }
            """;

        public ProjectAssistAgent(Kernel kernel, IServiceProvider service)
        {
            _kernel = kernel;
            _agent = new()
            {
                Instructions = AgentInstructions,
                Name = AgentName,
                Kernel = _kernel,
                Arguments = new KernelArguments(new OpenAIPromptExecutionSettings()
                {
                    FunctionChoiceBehavior = FunctionChoiceBehavior.Auto(),
                    ResponseFormat = "json_object"
                }),
            };
            _agent.Kernel.Plugins.Add(KernelPluginFactory.CreateFromType<PlannerPlugin>(serviceProvider: service));            
            _agent.Kernel.Plugins.Add(KernelPluginFactory.CreateFromType<AdaptiveCardPlugin>(serviceProvider: service));
            _agent.Kernel.Plugins.Add(KernelPluginFactory.CreateFromType<TeamsOutlookPlugin>(serviceProvider: service));
        }

        public async Task<ProjectAssistAgentResponse> InvokeAgentAsync(string input, ChatHistory chatHistory)
        {
            ArgumentNullException.ThrowIfNull(chatHistory);
            AgentThread thread = new ChatHistoryAgentThread();
            ChatMessageContent message = new(AuthorRole.User, input);
            chatHistory.Add(message);

            StringBuilder sb = new();
            await foreach (ChatMessageContent response in this._agent.InvokeAsync(chatHistory, thread: thread))
            {
                chatHistory.Add(response);
                sb.Append(response.Content);
            }

            try
            {
                string resultContent = sb.ToString();
                var jsonNode = JsonNode.Parse(resultContent);
                ProjectAssistAgentResponse result = new ProjectAssistAgentResponse()
                {
                    Content = jsonNode["content"].ToString(),
                    ContentType = Enum.Parse<ProjectAssistAgentResponseContentType>(jsonNode["contentType"].ToString(), true)
                };
                return result;
            }
            catch (Exception je)
            {
                return await InvokeAgentAsync($"That response did not match the expected format. Please try again. Error: {je.Message}", chatHistory);
            }
        }
    }
}

