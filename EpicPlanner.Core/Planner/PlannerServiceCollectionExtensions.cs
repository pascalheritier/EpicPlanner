using EpicPlanner.Core.Shared.Services;
using Microsoft.Extensions.DependencyInjection;

namespace EpicPlanner.Core.Planner;

public static class PlannerServiceCollectionExtensions
{
    public static IServiceCollection AddPlannerCore(this IServiceCollection _Services)
    {
        _Services.AddSingleton<PlanningDataProvider>();
        _Services.AddTransient<PlanningRunner>();
        return _Services;
    }
}
