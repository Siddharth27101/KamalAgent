namespace SprintReporting.Domain.Models;

public class AIInsightResult
{
    // ----- Narrative sections -----

    public string ExecutiveSummary { get; set; } = string.Empty;

    public string SprintOverview { get; set; } = string.Empty;

    public string StatusAnalysis { get; set; } = string.Empty;

    public string PriorityAnalysis { get; set; } = string.Empty;

    public string TeamWorkloadAnalysis { get; set; } = string.Empty;

    public string AssigneeProductivitySummary { get; set; } = string.Empty;

    public string ComponentAnalysis { get; set; } = string.Empty;

    public string LabelAnalysis { get; set; } = string.Empty;

    public string ResolutionSummary { get; set; } = string.Empty;

    // ----- Bulleted sections -----

    public List<string> Observations { get; set; }
        = new();

    public List<string> Risks { get; set; }
        = new();

    public List<string> Recommendations { get; set; }
        = new();

    public List<string> NextSprintSuggestions { get; set; }
        = new();

    // ----- Diagnostics -----

    public string ProviderUsed { get; set; } = "Fallback";

    public string DiagnosticMessage { get; set; } = string.Empty;
}
