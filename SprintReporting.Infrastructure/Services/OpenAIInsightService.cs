using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Options;
using SprintReporting.Application.Interfaces;
using SprintReporting.Domain.Models;
using SprintReporting.Infrastructure.Options;

namespace SprintReporting.Infrastructure.Services;

public class OpenAIInsightService : IAIInsightService
{
    private readonly HttpClient _httpClient;
    private readonly OpenAIOptions _options;

    public OpenAIInsightService(
        HttpClient httpClient,
        IOptions<OpenAIOptions> options)
    {
        _httpClient = httpClient;
        _options = options.Value;
    }

    public async Task<AIInsightResult> GenerateInsightsAsync(
        SprintMetrics metrics,
        ReportConfiguration configuration,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(_options.ApiKey))
        {
            return GenerateFallbackInsights(
                metrics,
                "OpenAI API key was not loaded at runtime.");
        }

        var prompt = BuildPrompt(metrics, configuration);

        var requestBody = new
        {
            model = _options.Model,
            temperature = 0.2,
            response_format = new
            {
                type = "json_object"
            },
            messages = new[]
            {
                new
                {
                    role = "system",
                    content = "You are an Agile Sprint Reporting Expert. Use only the provided aggregated metrics. Do not invent data. Return valid JSON only."
                },
                new
                {
                    role = "user",
                    content = prompt
                }
            }
        };

        using var request = new HttpRequestMessage(
            HttpMethod.Post,
            _options.Endpoint);

        request.Headers.Authorization = new AuthenticationHeaderValue(
            "Bearer",
            _options.ApiKey);

        request.Content = new StringContent(
            JsonSerializer.Serialize(requestBody),
            Encoding.UTF8,
            "application/json");

        HttpResponseMessage response;

        try
        {
            response = await _httpClient.SendAsync(
                request,
                cancellationToken);
        }
        catch (Exception ex)
        {
            return GenerateFallbackInsights(
                metrics,
                $"OpenAI HTTP request failed: {ex.GetType().Name}");
        }

        var responseContent = await response.Content.ReadAsStringAsync(
            cancellationToken);

        if (!response.IsSuccessStatusCode)
        {
            var safeError = BuildSafeErrorMessage(
                response.StatusCode.ToString(),
                responseContent);

            return GenerateFallbackInsights(
                metrics,
                safeError);
        }

        var aiText = ExtractAssistantText(responseContent);

        if (string.IsNullOrWhiteSpace(aiText))
        {
            return GenerateFallbackInsights(
                metrics,
                "OpenAI response did not contain assistant text.");
        }

        var parsedResult = TryParseInsightResult(aiText);

        if (parsedResult is not null)
        {
            parsedResult.ProviderUsed = "OpenAI";
            parsedResult.DiagnosticMessage = "OpenAI response parsed successfully.";
            return parsedResult;
        }

        return GenerateFallbackInsights(
            metrics,
            "OpenAI returned text, but JSON parsing failed.");
    }

    private static string BuildPrompt(
        SprintMetrics metrics,
        ReportConfiguration configuration)
    {
        var compactPayload = new
        {
            metrics.TotalIssues,
            metrics.CompletedIssues,
            metrics.OpenIssues,
            metrics.CompletionPercentage,
            metrics.HighPriorityOpenIssues,
            metrics.AverageIssueAgeDays,
            metrics.BacklogSize,
            metrics.TotalSprints,
            metrics.OverdueIssueCount,

            sprintIssueCount = metrics.SprintIssueCount,
            dueDateDistribution = metrics.DueDateDistribution,

            statusDistribution = metrics.StatusDistribution,
            priorityDistribution = metrics.PriorityDistribution,
            issueTypeDistribution = metrics.IssueTypeDistribution,
            resolutionDistribution = metrics.ResolutionDistribution,

            teamDistribution = metrics.TeamDistribution,
            reporterDistribution = metrics.ReporterDistribution,

            topAssignees = metrics.AssigneeDistribution
                .Take(8)
                .ToDictionary(),

            topContributors = metrics.TopContributors,

            topCompletedAssignees = metrics.CompletedWorkPerAssignee
                .Take(8)
                .ToDictionary(),

            topComponents = metrics.ComponentDistribution
                .Take(8)
                .ToDictionary(),

            labels = metrics.LabelDistribution,

            oldestOpenIssues = metrics.OldestOpenIssues
                .Take(5)
                .Select(issue => new
                {
                    issue.IssueKey,
                    issue.Priority,
                    issue.Status,
                    issue.AgeDays
                })
        };

        var jsonPayload = JsonSerializer.Serialize(
            compactPayload,
            new JsonSerializerOptions
            {
                WriteIndented = false
            });

        return $$"""
        Generate a sprint report using only this aggregated KPI data.

        Return JSON in this exact structure:
        {
          "executiveSummary": "",
          "sprintOverview": "",
          "statusAnalysis": "",
          "priorityAnalysis": "",
          "teamWorkloadAnalysis": "",
          "assigneeProductivitySummary": "",
          "componentAnalysis": "",
          "labelAnalysis": "",
          "resolutionSummary": "",
          "observations": ["", ""],
          "risks": ["", ""],
          "recommendations": ["", ""],
          "nextSprintSuggestions": ["", ""]
        }

        Rules:
        - Each narrative string field should be 2-4 concise sentences.
        - The list fields should each contain 3-5 short, concrete bullets.
        - Do not calculate new metrics; interpret only what is provided.
        - Do not mention raw Excel rows or invent data.
        - Use stakeholder-ready, professional language.
        - If a section has insufficient data, state that briefly instead of guessing.
        - Return valid JSON only, with no surrounding text.

        Aggregated KPI data:
        {{jsonPayload}}
        """;
    }

    private static string ExtractAssistantText(string responseContent)
    {
        try
        {
            using var document = JsonDocument.Parse(responseContent);

            var root = document.RootElement;

            if (!root.TryGetProperty("choices", out var choices))
            {
                return string.Empty;
            }

            if (choices.GetArrayLength() == 0)
            {
                return string.Empty;
            }

            var message = choices[0].GetProperty("message");

            return message.GetProperty("content").GetString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static AIInsightResult? TryParseInsightResult(string aiText)
    {
        try
        {
            var jsonText = ExtractJsonObject(aiText);

            if (string.IsNullOrWhiteSpace(jsonText))
            {
                return null;
            }

            var result = JsonSerializer.Deserialize<AIInsightResult>(
                jsonText,
                new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

            return result;
        }
        catch
        {
            return null;
        }
    }

    private static string ExtractJsonObject(string text)
    {
        var startIndex = text.IndexOf('{');
        var endIndex = text.LastIndexOf('}');

        if (startIndex < 0 || endIndex < 0 || endIndex <= startIndex)
        {
            return string.Empty;
        }

        return text.Substring(startIndex, endIndex - startIndex + 1);
    }

    private static string BuildSafeErrorMessage(
        string statusCode,
        string responseContent)
    {
        if (string.IsNullOrWhiteSpace(responseContent))
        {
            return $"OpenAI request failed with status code: {statusCode}.";
        }

        var compactResponse = responseContent.Length > 500
            ? responseContent.Substring(0, 500)
            : responseContent;

        return $"OpenAI request failed with status code: {statusCode}. Response: {compactResponse}";
    }

    private static AIInsightResult GenerateFallbackInsights(
        SprintMetrics metrics,
        string diagnosticMessage)
    {
        var completionMessage = metrics.CompletionPercentage >= 80
            ? "Sprint delivery is tracking strongly based on the current completion percentage."
            : "Sprint delivery needs attention because a significant portion of work remains open.";

        var riskMessage = metrics.HighPriorityOpenIssues > 0
            ? $"There are {metrics.HighPriorityOpenIssues} high-priority open issues that may affect delivery confidence."
            : "No high-priority open issues are currently visible in the calculated metrics.";

        var backlogMessage = metrics.BacklogSize > 0
            ? $"Backlog contains {metrics.BacklogSize} issues and should be reviewed for prioritization."
            : "Backlog volume appears low based on the calculated metrics.";

        var overdueMessage = metrics.OverdueIssueCount > 0
            ? $"{metrics.OverdueIssueCount} issues are past their due date and need immediate attention."
            : "No overdue issues were detected against the provided due dates.";

        var topStatus = DescribeTop(metrics.StatusDistribution);
        var topPriority = DescribeTop(metrics.PriorityDistribution);
        var topType = DescribeTop(metrics.IssueTypeDistribution);
        var topTeam = DescribeTop(metrics.TeamDistribution);
        var topAssignee = DescribeTop(metrics.AssigneeDistribution);
        var topComponent = DescribeTop(metrics.ComponentDistribution);
        var topLabel = DescribeTop(metrics.LabelDistribution);
        var topResolution = DescribeTop(metrics.ResolutionDistribution);

        return new AIInsightResult
        {
            ProviderUsed = "Fallback",
            DiagnosticMessage = diagnosticMessage,

            ExecutiveSummary =
                $"The sprint contains {metrics.TotalIssues} issues across {metrics.TotalSprints} sprint(s), with {metrics.CompletedIssues} completed and {metrics.OpenIssues} still open. Completion currently stands at {metrics.CompletionPercentage}%.",

            SprintOverview =
                $"Work is distributed across {metrics.TotalSprints} sprint(s) covering {metrics.TotalIssues} issues in total. {completionMessage}",

            StatusAnalysis =
                $"The most common status is {topStatus}. Open work totals {metrics.OpenIssues} issues against {metrics.CompletedIssues} completed.",

            PriorityAnalysis =
                $"The leading priority category is {topPriority}. {riskMessage}",

            TeamWorkloadAnalysis =
                $"The heaviest team allocation is {topTeam}. Assignee workload is led by {topAssignee}.",

            AssigneeProductivitySummary =
                $"{topAssignee} carries the most assigned work, while completed-work distribution highlights where delivery is concentrated. Average open issue age is {metrics.AverageIssueAgeDays} days.",

            ComponentAnalysis =
                $"Issue concentration is highest in component {topComponent}, which may warrant closer ownership and dependency review.",

            LabelAnalysis =
                $"The most frequent label is {topLabel}, giving a signal of the dominant work themes this sprint.",

            ResolutionSummary =
                $"The most common resolution outcome is {topResolution}. Issue type distribution is led by {topType}.",

            Observations = new List<string>
            {
                completionMessage,
                $"Average open issue age is {metrics.AverageIssueAgeDays} days.",
                $"Backlog currently holds {metrics.BacklogSize} issues."
            },

            Risks = new List<string>
            {
                riskMessage,
                backlogMessage,
                overdueMessage
            },

            Recommendations = new List<string>
            {
                "Review high-priority open items before stakeholder reporting.",
                "Use assignee and component distributions to rebalance workload concentration.",
                "Triage overdue and aging issues to protect the delivery timeline."
            },

            NextSprintSuggestions = new List<string>
            {
                "Carry over unresolved high-priority items with clear owners.",
                "Set realistic capacity based on the current completion rate.",
                "Break down aging issues into smaller, deliverable units."
            }
        };
    }

    private static string DescribeTop(Dictionary<string, int> distribution)
    {
        if (distribution is null || distribution.Count == 0)
        {
            return "not available";
        }

        var top = distribution
            .OrderByDescending(pair => pair.Value)
            .First();

        return $"{top.Key} ({top.Value})";
    }
}