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

        // Require either a thickness marker (t, thickness, dicke) or explicit mm unit.
        var matches = Regex.Matches(
            fileName,
            @"(?ix)
            (?:^|[^a-z0-9])
            (?:(?:t|thickness|dicke)\s*[:=_-]?\s*(?<value1>\d+(?:[\.,]\d+)?)|(?<value2>\d+(?:[\.,]\d+)?)\s*mm)
            (?:$|[^a-z0-9])");

        foreach (Match match in matches)
        {
            var raw = match.Groups["value1"].Success
                ? match.Groups["value1"].Value
                : match.Groups["value2"].Value;

            raw = raw.Replace(',', '.');
            if (!decimal.TryParse(raw, NumberStyles.Number, CultureInfo.InvariantCulture, out var thickness))
                continue;

            if (thickness <= 0m)
                continue;

            return thickness;
        }

        return null;
    }
}

