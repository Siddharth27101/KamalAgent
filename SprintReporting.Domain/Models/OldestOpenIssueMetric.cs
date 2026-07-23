namespace SprintReporting.Domain.Models;

public class OldestOpenIssueMetric
{
    public string IssueKey { get; set; } = string.Empty;

    public string Summary { get; set; } = string.Empty;

    public string Priority { get; set; } = string.Empty;

    public string Status { get; set; } = string.Empty;

    public int AgeDays { get; set; }
}