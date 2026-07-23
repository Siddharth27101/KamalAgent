using SprintReporting.Domain.Models;

namespace SprintReporting.Application.Interfaces;

public interface IAIInsightService
{
    Task<AIInsightResult> GenerateInsightsAsync(
        SprintMetrics metrics,
        ReportConfiguration configuration,
        CancellationToken cancellationToken = default);
}