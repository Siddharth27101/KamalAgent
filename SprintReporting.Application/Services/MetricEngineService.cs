using SprintReporting.Application.Interfaces;
using SprintReporting.Domain.Entities;
using SprintReporting.Domain.Models;

namespace SprintReporting.Application.Services;

public class MetricEngineService : IMetricEngineService
{
    public SprintMetrics CalculateMetrics(
        IReadOnlyList<IssueRecord> issues,
        ReportConfiguration configuration)
    {
        issues ??= Array.Empty<IssueRecord>();

        var totalIssues = issues.Count;

        var completedIssues = issues.Count(issue => IsDone(issue.Status));

        var openIssuesList = issues
            .Where(issue => !IsDone(issue.Status))
            .ToList();

        var openIssues = openIssuesList.Count;

        var highPriorityOpenIssues = openIssuesList.Count(issue =>
            IsHighPriority(issue.Priority));

        var backlogSize = issues.Count(issue =>
            IsBacklog(issue.Status));

        var averageIssueAgeDays = CalculateAverageIssueAgeDays(openIssuesList);

        var oldestOpenIssues = openIssuesList
            .Where(issue => issue.Created != default)
            .OrderBy(issue => issue.Created)
            .Take(5)
            .Select(issue => new OldestOpenIssueMetric
            {
                IssueKey = issue.IssueKey,
                Summary = issue.Summary,
                Priority = issue.Priority,
                Status = issue.Status,
                AgeDays = CalculateAgeDays(issue.Created)
            })
            .ToList();

        return new SprintMetrics
        {
            TotalIssues = totalIssues,
            CompletedIssues = completedIssues,
            OpenIssues = openIssues,
            CompletionPercentage = CalculateCompletionPercentage(
                completedIssues,
                totalIssues),
            HighPriorityOpenIssues = highPriorityOpenIssues,
            AverageIssueAgeDays = averageIssueAgeDays,
            BacklogSize = backlogSize,

            StatusDistribution = BuildDistribution(
                issues.Select(issue => issue.Status)),

            PriorityDistribution = BuildDistribution(
                issues.Select(issue => issue.Priority)),

            IssueTypeDistribution = BuildDistribution(
                issues.Select(issue => issue.IssueType)),

            AssigneeDistribution = BuildDistribution(
                issues.Select(issue => issue.Assignee)),

            CompletedWorkPerAssignee = BuildDistribution(
                issues
                    .Where(issue => IsDone(issue.Status))
                    .Select(issue => issue.Assignee)),

            ComponentDistribution = BuildMultiValueDistribution(
                issues.Select(issue => issue.Components)),

            LabelDistribution = BuildMultiValueDistribution(
                issues.Select(issue => issue.Labels)),

            OldestOpenIssues = oldestOpenIssues
        };
    }

    private static double CalculateCompletionPercentage(
        int completedIssues,
        int totalIssues)
    {
        if (totalIssues == 0)
        {
            return 0;
        }

        return Math.Round((double)completedIssues / totalIssues * 100, 2);
    }

    private static double CalculateAverageIssueAgeDays(
        IReadOnlyList<IssueRecord> openIssues)
    {
        var validIssues = openIssues
            .Where(issue => issue.Created != default)
            .ToList();

        if (validIssues.Count == 0)
        {
            return 0;
        }

        var average = validIssues
            .Average(issue => CalculateAgeDays(issue.Created));

        return Math.Round(average, 2);
    }

    private static int CalculateAgeDays(DateTime createdDate)
    {
        var today = DateTime.UtcNow.Date;
        var created = createdDate.Date;

        if (created > today)
        {
            return 0;
        }

        return (today - created).Days;
    }

    private static bool IsDone(string status)
    {
        return Normalize(status) == "done";
    }

    private static bool IsBacklog(string status)
    {
        return Normalize(status) == "backlog";
    }

    private static bool IsHighPriority(string priority)
    {
        var normalizedPriority = Normalize(priority);

        return normalizedPriority == "high"
            || normalizedPriority == "highest";
    }

    private static Dictionary<string, int> BuildDistribution(
        IEnumerable<string> values)
    {
        return values
            .Select(NormalizeDisplayValue)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .GroupBy(value => value, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key)
            .ToDictionary(
                group => group.Key,
                group => group.Count(),
                StringComparer.OrdinalIgnoreCase);
    }

    private static Dictionary<string, int> BuildMultiValueDistribution(
        IEnumerable<string> values)
    {
        return values
            .SelectMany(SplitMultiValue)
            .Select(NormalizeDisplayValue)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .GroupBy(value => value, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key)
            .ToDictionary(
                group => group.Key,
                group => group.Count(),
                StringComparer.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> SplitMultiValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return Array.Empty<string>();
        }

        return value.Split(
            new[] { ',', ';', '|' },
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
    }

    private static string Normalize(string value)
    {
        return value.Trim().ToLowerInvariant();
    }

    private static string NormalizeDisplayValue(string value)
    {
        return value.Trim();
    }
}