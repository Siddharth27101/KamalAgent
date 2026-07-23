using SprintReporting.Domain.Enums;

namespace SprintReporting.Application.DTOs;

public class GenerateReportRequestDto
{
    public List<ReportGroupType> SelectedGroups { get; set; } = new();
}