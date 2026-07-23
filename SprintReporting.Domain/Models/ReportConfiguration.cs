using SprintReporting.Domain.Enums;

namespace SprintReporting.Domain.Models;

public class ReportConfiguration
{
    public List<ReportGroupType> SelectedGroups { get; set; }
        = new();
}