using EpicPlanner.Core.Shared.Services;
using Microsoft.Extensions.DependencyInjection;

namespace EpicPlanner.Core.Planner;

public static class PlannerServiceCollectionExtensions
{
    public static IServiceCollection AddPlannerCore(this IServiceCollection services)
    {
        services.AddSingleton<PlanningDataProvider>();
        services.AddTransient<PlanningRunner>();
        return services;
    }
}
