namespace O365C.Agent.ProjectAssist
{
    public class ConfigOptions
    {
        public AzureConfigOptions Azure { get; set; }
        public GraphConfigOptions MicrosoftGraph { get; set; }
    }

    /// <summary>
    /// Options for Azure OpenAI and Azure Content Safety
    /// </summary>
    public class AzureConfigOptions
    {
        public string OpenAIApiKey { get; set; }
        public string OpenAIEndpoint { get; set; }
        public string OpenAIDeploymentName { get; set; }
    }
    public class GraphConfigOptions 
    {
        public string ClientId { get; set; }
        public string TenantId { get; set; }
        public string ClientSecret { get; set; }        

    }
}