using BendChecker.Core.Models;

namespace BendChecker.Core.Services;

public interface IStepAnalyzer
{
    Task<bool> CanOpenAsync(string stepPath, CancellationToken ct);
    Task<decimal?> TryGetThicknessMmAsync(string stepPath, CancellationToken ct);
    Task<StepScene?> TryLoadSceneAsync(string stepPath, CancellationToken ct);
}

