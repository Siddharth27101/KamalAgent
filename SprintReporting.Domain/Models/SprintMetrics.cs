namespace SprintReporting.Domain.Models;

public class SprintMetrics
{
    public int TotalIssues { get; set; }

    public int CompletedIssues { get; set; }

    public int OpenIssues { get; set; }

    public double CompletionPercentage { get; set; }

    public int HighPriorityOpenIssues { get; set; }

    public double AverageIssueAgeDays { get; set; }

    public int BacklogSize { get; set; }

    public Dictionary<string, int> StatusDistribution { get; set; }
        = new();

    public Dictionary<string, int> PriorityDistribution { get; set; }
        = new();

    public Dictionary<string, int> IssueTypeDistribution { get; set; }
        = new();

    public Dictionary<string, int> AssigneeDistribution { get; set; }
        = new();

    public Dictionary<string, int> CompletedWorkPerAssignee { get; set; }
        = new();

    public Dictionary<string, int> ComponentDistribution { get; set; }
        = new();

    public Dictionary<string, int> LabelDistribution { get; set; }
        = new();

    public List<OldestOpenIssueMetric> OldestOpenIssues { get; set; }
        = new();
}