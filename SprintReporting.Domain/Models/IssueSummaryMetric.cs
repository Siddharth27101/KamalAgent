namespace SprintReporting.Domain.Models;

/// <summary>
/// Lightweight projection of an issue used for timeline-style lists
/// (recently created / updated, overdue, etc.).
/// </summary>
public class IssueSummaryMetric
{
    public string IssueKey { get; set; } = string.Empty;

    public string Summary { get; set; } = string.Empty;

    public string Status { get; set; } = string.Empty;

    public string Assignee { get; set; } = string.Empty;

    public string Priority { get; set; } = string.Empty;

    /// <summary>
    /// The date relevant to the list context (created, updated, or due).
    /// </summary>
    public DateTime? Date { get; set; }

    /// <summary>
    /// Number of days the issue is overdue. Null when not applicable.
    /// </summary>
    public int? DaysOverdue { get; set; }
}
