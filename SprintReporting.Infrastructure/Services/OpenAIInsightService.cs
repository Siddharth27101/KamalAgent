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
            selectedGroups = configuration.SelectedGroups
                .Select(group => group.ToString())
                .ToList(),

            metrics.TotalIssues,
            metrics.CompletedIssues,
            metrics.OpenIssues,
            metrics.CompletionPercentage,
            metrics.HighPriorityOpenIssues,
            metrics.AverageIssueAgeDays,
            metrics.BacklogSize,

            statusDistribution = metrics.StatusDistribution,
            priorityDistribution = metrics.PriorityDistribution,
            issueTypeDistribution = metrics.IssueTypeDistribution,

            topAssignees = metrics.AssigneeDistribution
                .Take(8)
                .ToDictionary(),

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
        Generate sprint report insights using only this aggregated KPI data.

        Return JSON in this exact structure:
        {
          "executiveSummary": "",
          "observations": ["", ""],
          "risks": ["", ""],
          "recommendations": ["", ""]
        }

        Rules:
        - Do not calculate new metrics.
        - Do not mention raw Excel rows.
        - Keep each bullet concise.
        - Focus on stakeholder-ready language.
        - If data is insufficient, say so clearly.
        - Return JSON only.

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

        return new AIInsightResult
        {
            ProviderUsed = "Fallback",
            DiagnosticMessage = diagnosticMessage,

            ExecutiveSummary =
                $"The sprint contains {metrics.TotalIssues} issues, with {metrics.CompletedIssues} completed and {metrics.OpenIssues} still open. Completion currently stands at {metrics.CompletionPercentage}%.",

            Observations = new List<string>
            {
                completionMessage,
                $"Average open issue age is {metrics.AverageIssueAgeDays} days."
            },

            Risks = new List<string>
            {
                riskMessage,
                backlogMessage
            },

            Recommendations = new List<string>
            {
                "Review high-priority open items before stakeholder reporting.",
                "Use assignee and component distributions to identify workload concentration."
            }
        };
    }
}