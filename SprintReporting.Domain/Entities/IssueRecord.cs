namespace SprintReporting.Domain.Entities;

public class IssueRecord
{
    public string IssueType { get; set; } = string.Empty;

    public string IssueKey { get; set; } = string.Empty;

    public long IssueId { get; set; }

    public string Summary { get; set; } = string.Empty;

    public string Sprint { get; set; } = string.Empty;

    public string Assignee { get; set; } = string.Empty;

    public string Reporter { get; set; } = string.Empty;

    public string Team { get; set; } = string.Empty;

    public string Priority { get; set; } = string.Empty;

    public string Status { get; set; } = string.Empty;

    public string Resolution { get; set; } = string.Empty;

    public DateTime Created { get; set; }

    public DateTime Updated { get; set; }

    public DateTime? DueDate { get; set; }

    public string Labels { get; set; } = string.Empty;

    public string Components { get; set; } = string.Empty;
}