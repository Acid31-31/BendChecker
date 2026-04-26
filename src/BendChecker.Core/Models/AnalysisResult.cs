namespace BendChecker.Core.Models;

public sealed record AnalysisResult(
    string StepPath,
    string RulesPath,
    string Material,
    decimal ThicknessMm,
    RuleRow? SelectedRule,
    IReadOnlyList<Finding> Findings
);
