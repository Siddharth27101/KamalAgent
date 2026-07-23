using SprintReporting.Domain.Entities;
using SprintReporting.Domain.Models;

namespace SprintReporting.Application.Interfaces;

public interface IMetricEngineService
{
    SprintMetrics CalculateMetrics(
        IReadOnlyList<IssueRecord> issues,
        ReportConfiguration configuration);
}