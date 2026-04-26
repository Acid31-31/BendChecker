using System.Globalization;
using System.Text.RegularExpressions;

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

    public Task<decimal?> TryGetThicknessMmAsync(string stepPath, CancellationToken ct)
    {
        return Task.Run<decimal?>(() =>
        {
            ct.ThrowIfCancellationRequested();

            if (!File.Exists(stepPath))
                return null;

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
        }, ct);
    }
}

