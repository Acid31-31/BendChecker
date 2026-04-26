using BendChecker.Core.Models;

namespace BendChecker.Core.Services;

public interface IStepPreviewLoader
{
    Task<StepPreviewScene> LoadPreviewAsync(string stepPath, CancellationToken ct);
}
