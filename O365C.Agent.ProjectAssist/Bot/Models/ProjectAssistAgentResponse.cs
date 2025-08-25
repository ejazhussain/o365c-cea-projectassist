using System.ComponentModel;
using System.Text.Json.Serialization;

namespace O365C.Agent.ProjectAssist.Bot.Models;

public enum ProjectAssistAgentResponseContentType
{
    [JsonPropertyName("text")]
    Text,

    [JsonPropertyName("adaptive-card")]
    AdaptiveCard
}

public class ProjectAssistAgentResponse
{
    [JsonPropertyName("contentType")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ProjectAssistAgentResponseContentType ContentType { get; set; }

    [JsonPropertyName("content")]
    [Description("The content of the response, may be plain text, or JSON based adaptive card but must be a string.")]
    public string Content { get; set; }
}
