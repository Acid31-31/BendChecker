using System.Globalization;
using System.Text.RegularExpressions;
using BendChecker.Core.Models;

namespace BendChecker.Core.Services;

public sealed class StepAnalyzerStub : IStepAnalyzer
{
    private static readonly StepScene SampleScene = new([
        new StepMeshPart(
            Positions:
            [
                -20d, -10d, -2d,
                20d, -10d, -2d,
                20d, 10d, -2d,
                -20d, -10d, -2d,
                20d, 10d, -2d,
                -20d, 10d, -2d,
                -20d, -10d, 2d,
                20d, 10d, 2d,
                20d, -10d, 2d,
                -20d, -10d, 2d,
                -20d, 10d, 2d,
                20d, 10d, 2d
            ],
            Normals:
            [
                0d, 0d, -1d,
                0d, 0d, -1d,
                0d, 0d, -1d,
                0d, 0d, -1d,
                0d, 0d, -1d,
                0d, 0d, -1d,
                0d, 0d, 1d,
                0d, 0d, 1d,
                0d, 0d, 1d,
                0d, 0d, 1d,
                0d, 0d, 1d,
                0d, 0d, 1d
            ],
            Indices: Enumerable.Range(0, 12).ToArray(),
            Red: 185,
            Green: 190,
            Blue: 205,
            Alpha: 255)
    ]);

    public Task<bool> CanOpenAsync(string stepPath, CancellationToken ct)
    {
        var ok = File.Exists(stepPath) &&
                 (stepPath.EndsWith(".step", StringComparison.OrdinalIgnoreCase) ||
                  stepPath.EndsWith(".stp", StringComparison.OrdinalIgnoreCase));
        return Task.FromResult(ok);
    }

    public Task<decimal?> TryGetThicknessMmAsync(string stepPath, CancellationToken ct)
    {
        return Task.Run<decimal?>(() =>
        {
            ct.ThrowIfCancellationRequested();

            if (!File.Exists(stepPath))
                return null;

            return TryParseThicknessFromName(stepPath);
        }, ct);
    }

    public Task<StepScene?> TryLoadSceneAsync(string stepPath, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        return Task.FromResult<StepScene?>(SampleScene);
    }

    internal static decimal? TryParseThicknessFromName(string stepPath)
    {
        var fileName = Path.GetFileNameWithoutExtension(stepPath);
        var matches = Regex.Matches(
            fileName,
            @"(?ix)
            (?:^|[^0-9])
            (?:t|thickness|dicke)?
            \s*[:=_-]?
            (?<value>\d+(?:[\.,]\d+)?)
            \s*(?:mm)?
            (?:$|[^0-9])");

        foreach (Match match in matches)
        {
            var raw = match.Groups["value"].Value.Replace(',', '.');
            if (decimal.TryParse(raw, NumberStyles.Number, CultureInfo.InvariantCulture, out var thickness))
                return thickness;
        }

        return null;
    }
}

