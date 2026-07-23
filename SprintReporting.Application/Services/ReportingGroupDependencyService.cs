using SprintReporting.Application.Interfaces;
using SprintReporting.Domain.Enums;
using SprintReporting.Domain.Models;

namespace SprintReporting.Application.Services;

public class ReportingGroupDependencyService : IReportingGroupDependencyService
{
    public ReportConfiguration BuildConfiguration(
        IReadOnlyList<ReportGroupType> selectedGroups,
        bool includeAllGroups = false)
    {
        // "All" selection: include every report group in its canonical order.
        if (includeAllGroups)
        {
            return new ReportConfiguration
            {
                SelectedGroups = Enum.GetValues<ReportGroupType>().ToList()
            };
        }

        var normalizedGroups = selectedGroups
            .Where(group => Enum.IsDefined(typeof(ReportGroupType), group))
            .Distinct()
            .ToList();

        if (normalizedGroups.Count == 0)
        {
            normalizedGroups.Add(ReportGroupType.Delivery);
            normalizedGroups.Add(ReportGroupType.PriorityRisk);
            normalizedGroups.Add(ReportGroupType.TeamAnalysis);
        }

        return new ReportConfiguration
        {
            SelectedGroups = normalizedGroups
        };
    }
}