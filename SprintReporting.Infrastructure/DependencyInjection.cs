using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using SprintReporting.Application.Interfaces;
using SprintReporting.Infrastructure.Options;
using SprintReporting.Infrastructure.Services;

namespace SprintReporting.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructureServices(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        services.Configure<OpenAIOptions>(
            configuration.GetSection("OpenAI"));

        services.AddScoped<IExcelParserService, ExcelParserService>();

        services.AddScoped<IPowerPointService, PowerPointService>();

        services.AddHttpClient<IAIInsightService, OpenAIInsightService>();

        return services;
    }
}