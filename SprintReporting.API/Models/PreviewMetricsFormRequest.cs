using Microsoft.AspNetCore.Http;
using SprintReporting.Domain.Enums;

namespace SprintReporting.API.Models;

public class PreviewMetricsFormRequest
{
    public IFormFile? File { get; set; }

    public List<ReportGroupType> SelectedGroups { get; set; } = new();

    /// <summary>
    /// When true, every report group is included (the "All" option),
    /// regardless of what is passed in <see cref="SelectedGroups"/>.
    /// </summary>
    public bool IncludeAllGroups { get; set; }
}