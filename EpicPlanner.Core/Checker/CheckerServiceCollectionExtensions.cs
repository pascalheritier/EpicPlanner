using EpicPlanner.Core.Shared.Services;
using Microsoft.Extensions.DependencyInjection;

namespace EpicPlanner.Core.Checker;

public static class CheckerServiceCollectionExtensions
{
    public static IServiceCollection AddCheckerCore(this IServiceCollection _Services)
    {
        _Services.AddSingleton<PlanningDataProvider>();
        _Services.AddTransient<CheckingRunner>();
        return _Services;
    }
}
