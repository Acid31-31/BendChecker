namespace BendChecker.Core.Services;

public sealed class StepAnalyzerStub : IStepAnalyzer
{
    public Task<bool> CanOpenAsync(string stepPath, CancellationToken ct)
    {
        var ok = File.Exists(stepPath) &&
                 (stepPath.EndsWith(".step", StringComparison.OrdinalIgnoreCase) ||
                  stepPath.EndsWith(".stp", StringComparison.OrdinalIgnoreCase));
        return Task.FromResult(ok);
    }
}
