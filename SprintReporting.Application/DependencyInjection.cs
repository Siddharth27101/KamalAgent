using Microsoft.Extensions.DependencyInjection;
using SprintReporting.Application.Interfaces;
using SprintReporting.Application.Services;

namespace SprintReporting.Application;

public static class DependencyInjection
{
    public static IServiceCollection AddApplicationServices(
        this IServiceCollection services)
    {
        services.AddScoped<IMetricEngineService, MetricEngineService>();
        services.AddScoped<IReportingGroupDependencyService, ReportingGroupDependencyService>();

        return services;
    }
}
