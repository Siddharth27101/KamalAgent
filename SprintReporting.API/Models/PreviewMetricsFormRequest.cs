using Microsoft.AspNetCore.Http;
using SprintReporting.Domain.Enums;

namespace SprintReporting.API.Models;

public class PreviewMetricsFormRequest
{
    public IFormFile? File { get; set; }

    public List<ReportGroupType> SelectedGroups { get; set; } = new();
}