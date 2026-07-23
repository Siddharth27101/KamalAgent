using SprintReporting.Domain.Models;

namespace SprintReporting.Application.Interfaces;

public interface IPowerPointService
{
    Task<byte[]> GeneratePresentationAsync(
        SprintMetrics metrics,
        AIInsightResult aiInsights,
        ReportConfiguration configuration,
        CancellationToken cancellationToken = default);
}