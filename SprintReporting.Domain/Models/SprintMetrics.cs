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

    // ----- Sprint overview -----

    public int TotalSprints { get; set; }

    public Dictionary<string, int> SprintIssueCount { get; set; }
        = new();

    // ----- Timeline analysis -----

    public int OverdueIssueCount { get; set; }

    public Dictionary<string, int> DueDateDistribution { get; set; }
        = new();

    public List<IssueSummaryMetric> RecentlyCreatedIssues { get; set; }
        = new();

    public List<IssueSummaryMetric> RecentlyUpdatedIssues { get; set; }
        = new();

    public List<IssueSummaryMetric> OverdueIssues { get; set; }
        = new();

    // ----- Distributions -----

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

    public Dictionary<string, int> ResolutionDistribution { get; set; }
        = new();

    public Dictionary<string, int> TeamDistribution { get; set; }
        = new();

    public Dictionary<string, int> ReporterDistribution { get; set; }
        = new();

    /// <summary>
    /// Top assignees by number of assigned issues.
    /// </summary>
    public Dictionary<string, int> TopContributors { get; set; }
        = new();

    public List<OldestOpenIssueMetric> OldestOpenIssues { get; set; }
        = new();
}