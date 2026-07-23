using SprintReporting.Domain.Enums;
using SprintReporting.Domain.Models;

namespace SprintReporting.Application.Interfaces;

public interface IReportingGroupDependencyService
{
    ReportConfiguration BuildConfiguration(
        IReadOnlyList<ReportGroupType> selectedGroups,
        bool includeAllGroups = false);
}