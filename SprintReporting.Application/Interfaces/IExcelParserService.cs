using SprintReporting.Domain.Entities;

namespace SprintReporting.Application.Interfaces;

public interface IExcelParserService
{
    Task<IReadOnlyList<IssueRecord>> ParseAsync(
        Stream excelStream,
        CancellationToken cancellationToken = default);
}