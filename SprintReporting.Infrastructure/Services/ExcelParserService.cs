using System.Globalization;
using ClosedXML.Excel;
using SprintReporting.Application.Interfaces;
using SprintReporting.Domain.Entities;

namespace SprintReporting.Infrastructure.Services;

public class ExcelParserService : IExcelParserService
{
    public Task<IReadOnlyList<IssueRecord>> ParseAsync(
        Stream excelStream,
        CancellationToken cancellationToken = default)
    {
        if (excelStream is null)
        {
            throw new ArgumentNullException(nameof(excelStream));
        }

        cancellationToken.ThrowIfCancellationRequested();

        using var workbook = new XLWorkbook(excelStream);

        var worksheet = workbook.Worksheets.FirstOrDefault();

        if (worksheet is null)
        {
            throw new InvalidOperationException("The uploaded Excel file does not contain any worksheet.");
        }

        var headerRow = worksheet.FirstRowUsed();

        if (headerRow is null)
        {
            throw new InvalidOperationException("The uploaded Excel file is empty.");
        }

        var headers = BuildHeaderMap(headerRow);

        var issueRecords = new List<IssueRecord>();

        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var issueKey = GetString(row, headers, "Issue key");

            if (string.IsNullOrWhiteSpace(issueKey))
            {
                continue;
            }

            var issueRecord = new IssueRecord
            {
                IssueType = GetString(row, headers, "Issue Type"),
                IssueKey = issueKey,
                IssueId = GetLong(row, headers, "Issue id"),
                Summary = GetString(row, headers, "Summary"),
                Sprint = GetString(row, headers, "Sprint"),
                Assignee = GetString(row, headers, "Assignee"),
                Reporter = GetString(row, headers, "Reporter"),
                Team = GetString(row, headers, "Team"),
                Priority = GetString(row, headers, "Priority"),
                Status = GetString(row, headers, "Status"),
                Resolution = GetString(row, headers, "Resolution"),
                Created = GetDate(row, headers, "Created"),
                Updated = GetDate(row, headers, "Updated"),
                DueDate = GetNullableDate(row, headers, "Due date"),
                Labels = GetString(row, headers, "Labels"),
                Components = GetString(row, headers, "Components")
            };

            issueRecords.Add(issueRecord);
        }

        return Task.FromResult<IReadOnlyList<IssueRecord>>(issueRecords);
    }

    private static Dictionary<string, int> BuildHeaderMap(IXLRow headerRow)
    {
        var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var cell in headerRow.CellsUsed())
        {
            var headerName = NormalizeHeader(cell.GetFormattedString());

            if (string.IsNullOrWhiteSpace(headerName))
            {
                continue;
            }

            if (!headers.ContainsKey(headerName))
            {
                headers.Add(headerName, cell.Address.ColumnNumber);
            }
        }

        return headers;
    }

    private static string GetString(
        IXLRow row,
        Dictionary<string, int> headers,
        string columnName)
    {
        if (!TryGetColumnNumber(headers, columnName, out var columnNumber))
        {
            return string.Empty;
        }

        var cell = row.Cell(columnNumber);

        if (cell.IsEmpty())
        {
            return string.Empty;
        }

        return cell.GetFormattedString().Trim();
    }

    private static long GetLong(
        IXLRow row,
        Dictionary<string, int> headers,
        string columnName)
    {
        if (!TryGetColumnNumber(headers, columnName, out var columnNumber))
        {
            return 0;
        }

        var cell = row.Cell(columnNumber);

        if (cell.IsEmpty())
        {
            return 0;
        }

        if (cell.TryGetValue<long>(out var longValue))
        {
            return longValue;
        }

        var textValue = cell.GetFormattedString().Trim();

        if (long.TryParse(textValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedLong))
        {
            return parsedLong;
        }

        if (decimal.TryParse(textValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedDecimal))
        {
            return Convert.ToInt64(parsedDecimal);
        }

        return 0;
    }

    private static DateTime GetDate(
        IXLRow row,
        Dictionary<string, int> headers,
        string columnName)
    {
        return GetNullableDate(row, headers, columnName) ?? default;
    }

    private static DateTime? GetNullableDate(
        IXLRow row,
        Dictionary<string, int> headers,
        string columnName)
    {
        if (!TryGetColumnNumber(headers, columnName, out var columnNumber))
        {
            return null;
        }

        var cell = row.Cell(columnNumber);

        if (cell.IsEmpty())
        {
            return null;
        }

        if (cell.TryGetValue<DateTime>(out var dateValue))
        {
            return dateValue;
        }

        var textValue = cell.GetFormattedString().Trim();

        if (DateTime.TryParse(
                textValue,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeLocal,
                out var parsedDate))
        {
            return parsedDate;
        }

        if (DateTime.TryParse(
                textValue,
                CultureInfo.CurrentCulture,
                DateTimeStyles.AssumeLocal,
                out parsedDate))
        {
            return parsedDate;
        }
        return null;
    }

    private static bool TryGetColumnNumber(
        Dictionary<string, int> headers,
        string columnName,
        out int columnNumber)
    {
        var normalizedColumnName = NormalizeHeader(columnName);

        return headers.TryGetValue(normalizedColumnName, out columnNumber);
    }

    private static string NormalizeHeader(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return value
            .Trim()
            .Replace(" ", string.Empty)
            .Replace("_", string.Empty)
            .Replace("-", string.Empty)
            .ToLowerInvariant();
    }
}