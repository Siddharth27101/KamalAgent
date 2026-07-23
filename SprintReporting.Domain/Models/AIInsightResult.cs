namespace SprintReporting.Domain.Models;

public class AIInsightResult
{
    public string ExecutiveSummary { get; set; } = string.Empty;

    public List<string> Observations { get; set; }
        = new();

    public List<string> Risks { get; set; }
        = new();

    public List<string> Recommendations { get; set; }
        = new();

    public string ProviderUsed { get; set; } = "Fallback";

    public string DiagnosticMessage { get; set; } = string.Empty;
}