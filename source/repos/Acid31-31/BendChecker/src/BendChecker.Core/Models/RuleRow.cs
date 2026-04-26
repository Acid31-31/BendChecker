namespace BendChecker.Core.Models;

public sealed record RuleRow(
    string Material,
    decimal ThicknessMm,
    string PrismaV,
    string? Zuschlagverfahren,
    string? Biegeradien,
    decimal? Massabzug,
    decimal? Sollmass90,
    decimal? Abwicklungsmaß90,
    decimal? MinSchenkelMm
);
