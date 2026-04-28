using BendChecker.Core.Models;

namespace BendChecker.Core.Services;

public sealed class BendCheckService(RuleService ruleService, IStepAnalyzer stepAnalyzer)
{
    public async Task<AnalysisResult> AnalyzeAsync(
        string stepPath,
        string rulesXlsxPath,
        string material,
        decimal thicknessMm,
        string? prismaV,
        CancellationToken ct)
    {
        var rules = ruleService.LoadRules(rulesXlsxPath);
        return await AnalyzeCoreAsync(stepPath, rulesXlsxPath, rules, material, thicknessMm, prismaV, ct);
    }

    public async Task<AnalysisResult> AnalyzeAsync(
        string stepPath,
        IReadOnlyList<RuleRow> rules,
        string material,
        decimal thicknessMm,
        string? prismaV,
        CancellationToken ct)
    {
        return await AnalyzeCoreAsync(stepPath, "ImportedRulesDb", rules, material, thicknessMm, prismaV, ct);
    }

    private async Task<AnalysisResult> AnalyzeCoreAsync(
        string stepPath,
        string rulesPath,
        IReadOnlyList<RuleRow> rules,
        string material,
        decimal thicknessMm,
        string? prismaV,
        CancellationToken ct)
    {
        var findings = new List<Finding>();

        var canOpen = await stepAnalyzer.CanOpenAsync(stepPath, ct);
        if (!canOpen)
            findings.Add(new Finding("FAIL", "STEP_OPEN", "STEP-Datei kann nicht geöffnet werden (Stub)."));

        var rule = ruleService.FindBestRule(rules, material, thicknessMm, prismaV);

        if (rule is null)
        {
            findings.Add(new Finding("FAIL", "RULE_NOT_FOUND", "Keine passende Regelzeile fuer Material/Dicke/V gefunden."));
        }
        else
        {
            if (rule.MinSchenkelMm is null)
                findings.Add(new Finding("WARN", "MIN_FLANGE_EMPTY", "Schenkelmas minimal ist in der Regelzeile leer."));
            else
                findings.Add(new Finding("OK", "MIN_FLANGE_RULE",
                    $"Schenkelmas minimal laut Tabelle: {rule.MinSchenkelMm:0.##} mm (Messung kommt mit STEP-Engine)."));

            findings.Add(new Finding("OK", "TOOL_DEFAULT", "Standardwerkzeug angenommen: 90 Matrize, 88 Stempel, Spitzenradius ~1mm."));
            findings.Add(new Finding("INFO", "AUTO_SUGGEST", "Auto-Verbesserung V1: Vorschlaege textlich; echte STEP-Geometrieaenderungen folgen spaeter."));
        }

        return new AnalysisResult(stepPath, rulesPath, material, thicknessMm, rule, findings);
    }
}

