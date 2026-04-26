namespace BendChecker.Core.Services;

public interface IStepAnalyzer
{
    Task<bool> CanOpenAsync(string stepPath, CancellationToken ct);
}
