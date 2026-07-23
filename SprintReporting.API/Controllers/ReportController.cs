using Microsoft.AspNetCore.Mvc;
using SprintReporting.API.Models;
using SprintReporting.Application.Interfaces;
using SprintReporting.Domain.Enums;

namespace SprintReporting.API.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ReportsController : ControllerBase
{
    private readonly IExcelParserService _excelParserService;
    private readonly IMetricEngineService _metricEngineService;
    private readonly IReportingGroupDependencyService _reportingGroupDependencyService;
    private readonly IAIInsightService _aiInsightService;
    private readonly IPowerPointService _powerPointService;

    public ReportsController(
        IExcelParserService excelParserService,
        IMetricEngineService metricEngineService,
        IReportingGroupDependencyService reportingGroupDependencyService,
        IAIInsightService aiInsightService,
        IPowerPointService powerPointService)
    {
        _excelParserService = excelParserService;
        _metricEngineService = metricEngineService;
        _reportingGroupDependencyService = reportingGroupDependencyService;
        _aiInsightService = aiInsightService;
        _powerPointService = powerPointService;
    }

    [HttpPost("preview-metrics")]
    [Consumes("multipart/form-data")]
    [RequestSizeLimit(20_000_000)]
    public async Task<IActionResult> PreviewMetrics(
        [FromForm] PreviewMetricsFormRequest request,
        CancellationToken cancellationToken)
    {
        var validationResult = ValidateFile(request);

        if (validationResult is not null)
        {
            return validationResult;
        }

        var issues = await ParseIssuesAsync(request, cancellationToken);

        if (issues.Count == 0)
        {
            return BadRequest(new
            {
                message = "No valid issue records were found in the uploaded Excel file."
            });
        }

        var configuration = _reportingGroupDependencyService.BuildConfiguration(
            request.SelectedGroups ?? new List<ReportGroupType>());

        var metrics = _metricEngineService.CalculateMetrics(
            issues,
            configuration);

        return Ok(new
        {
            message = "Metrics generated successfully.",
            selectedGroups = configuration.SelectedGroups,
            totalParsedIssues = issues.Count,
            metrics
        });
    }

    [HttpPost("preview-insights")]
    [Consumes("multipart/form-data")]
    [RequestSizeLimit(20_000_000)]
    public async Task<IActionResult> PreviewInsights(
        [FromForm] PreviewMetricsFormRequest request,
        CancellationToken cancellationToken)
    {
        var validationResult = ValidateFile(request);

        if (validationResult is not null)
        {
            return validationResult;
        }

        var issues = await ParseIssuesAsync(request, cancellationToken);

        if (issues.Count == 0)
        {
            return BadRequest(new
            {
                message = "No valid issue records were found in the uploaded Excel file."
            });
        }

        var configuration = _reportingGroupDependencyService.BuildConfiguration(
            request.SelectedGroups ?? new List<ReportGroupType>());

        var metrics = _metricEngineService.CalculateMetrics(
            issues,
            configuration);

        var aiInsights = await _aiInsightService.GenerateInsightsAsync(
            metrics,
            configuration,
            cancellationToken);

        return Ok(new
        {
            message = "AI insights generated successfully.",
            note = "If OpenAI ApiKey is empty, local fallback insights are returned.",
            selectedGroups = configuration.SelectedGroups,
            totalParsedIssues = issues.Count,
            metrics,
            aiInsights
        });
    }

    [HttpPost("generate-pptx")]
    [Consumes("multipart/form-data")]
    [RequestSizeLimit(20_000_000)]
    public async Task<IActionResult> GeneratePowerPoint(
        [FromForm] PreviewMetricsFormRequest request,
        CancellationToken cancellationToken)
    {
        var validationResult = ValidateFile(request);

        if (validationResult is not null)
        {
            return validationResult;
        }

        var issues = await ParseIssuesAsync(request, cancellationToken);

        if (issues.Count == 0)
        {
            return BadRequest(new
            {
                message = "No valid issue records were found in the uploaded Excel file."
            });
        }

        var configuration = _reportingGroupDependencyService.BuildConfiguration(
            request.SelectedGroups ?? new List<ReportGroupType>());

        var metrics = _metricEngineService.CalculateMetrics(
            issues,
            configuration);

        var aiInsights = await _aiInsightService.GenerateInsightsAsync(
            metrics,
            configuration,
            cancellationToken);

        var pptxBytes = await _powerPointService.GeneratePresentationAsync(
            metrics,
            aiInsights,
            configuration,
            cancellationToken);

        var fileName = $"Sprint_Report_{DateTime.Now:yyyyMMdd_HHmmss}.pptx";

        return File(
            pptxBytes,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            fileName);
    }

    private async Task<IReadOnlyList<SprintReporting.Domain.Entities.IssueRecord>> ParseIssuesAsync(
        PreviewMetricsFormRequest request,
        CancellationToken cancellationToken)
    {
        await using var stream = request.File!.OpenReadStream();

        return await _excelParserService.ParseAsync(
            stream,
            cancellationToken);
    }

    private static IActionResult? ValidateFile(
        PreviewMetricsFormRequest request)
    {
        if (request.File is null || request.File.Length == 0)
        {
            return new BadRequestObjectResult(new
            {
                message = "Please upload a valid Excel file."
            });
        }

        if (!IsExcelFile(request.File.FileName))
        {
            return new BadRequestObjectResult(new
            {
                message = "Only .xlsx Excel files are supported."
            });
        }

        return null;
    }

    private static bool IsExcelFile(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return false;
        }

        return Path.GetExtension(fileName)
            .Equals(".xlsx", StringComparison.OrdinalIgnoreCase);
    }
}