using EpicPlanner.Core.Shared.Services;
using Microsoft.Extensions.DependencyInjection;

namespace EpicPlanner.Core.Checker;

public static class CheckerServiceCollectionExtensions
{
    public static IServiceCollection AddCheckerCore(this IServiceCollection services)
    {
        services.AddSingleton<PlanningDataProvider>();
        services.AddTransient<CheckingRunner>();
        return services;
    }
}
