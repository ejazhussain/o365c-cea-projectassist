using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.Graph.Models;
using System.Text.Json.Serialization;

namespace O365C.Agent.ProjectAssist.Bot.Models
{
    public class PlannerTaskResponse
    {
        public string Title { get; set; }
        public string Id { get; set; }
        public string PlanId { get; set; }
        public string BucketId { get; set; }
        public string OrderHint { get; set; }
        public string AssigneePriority { get; set; }
        public DateTimeOffset CreatedDateTime { get; set; }
        public IdentitySet CreatedBy { get; set; }
        public IDictionary<string, PlannerAssignment> Assignments { get; set; }

        public bool HasDescription { get; set; }
        public string Description
        {
            get => HasDescription && Details != null ? Details.Description : null;
            set
            {
                if (Details == null)
                {
                    Details = new PlannerTaskDetails();
                }
                Details.Description = value;
            }
        }
        public PlannerTaskDetails Details { get; set; }
        public int PercentComplete { get; set; }
        public int Priority { get; set; }
        public string PriorityLabel
        {
            get
            {
                if (Priority == 0 || Priority == 1)
                    return "Urgent";
                if (Priority == 2 || Priority == 3 || Priority == 4)
                    return "Important";
                if (Priority == 5 || Priority == 6 || Priority == 7)
                    return "Medium";
                if (Priority == 8 || Priority == 9 || Priority == 10)
                    return "Low";
                return "Unknown";
            }
        }
        public string ProgressLabel
        {
            get
            {
                if (PercentComplete >= 100)
                    return "Completed";
                if (PercentComplete > 0)
                    return "In Progress";
                return "Not started";
            }
        }
        public DateTimeOffset? StartDateTime { get; set; }
        public DateTimeOffset? DueDateTime { get; set; }
    }   
}
