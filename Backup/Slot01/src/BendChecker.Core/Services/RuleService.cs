using BendChecker.Core.Models;
using ClosedXML.Excel;

namespace BendChecker.Core.Services;

public sealed class RuleService
{
    public List<RuleRow> LoadRules(string xlsxPath, string? sheetName = null)
    {
        using var wb = new XLWorkbook(xlsxPath);
        var ws = sheetName is null ? wb.Worksheets.First() : wb.Worksheet(sheetName);
        return LoadRulesFromWorksheet(ws);
    }

    public List<RuleRow> LoadRulesFromAllSheets(string xlsxPath)
    {
        using var wb = new XLWorkbook(xlsxPath);
        var all = new List<RuleRow>();
        foreach (var ws in wb.Worksheets)
        {
            try
            {
                all.AddRange(LoadRulesFromWorksheet(ws));
            }
            catch
            {
                // skip sheets that do not match expected bending rule format
            }
        }

        return all;
    }

    private static List<RuleRow> LoadRulesFromWorksheet(IXLWorksheet ws)
    {
        var headerRow = ws.FirstRowUsed() ?? throw new InvalidOperationException("Excel enthält keine Kopfzeile.");
        var headerMap = headerRow.CellsUsed()
            .ToDictionary(c => c.GetString().Trim(), c => c.Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);

        int Col(params string[] names)
        {
            foreach (var name in names)
            {
                if (headerMap.TryGetValue(name, out var col))
                    return col;
            }

            throw new InvalidOperationException($"Excel-Spalte fehlt: '{string.Join("' / '", names)}'. Vorhanden: {string.Join(", ", headerMap.Keys)}");
        }

        var cMaterial = Col("Material");
        var cT = Col("Materialstärke", "Materialstaerke");
        var cV = Col("Prisma V");
        var cZ = Col("Zuschlagverfahren");
        var cR = Col("Biegeradien");
        var cAbzug = Col("Maßabzug", "Massabzug");
        var cSoll = Col("Sollmaß 90°", "Sollmaß 90", "Sollmass 90°", "Sollmass 90");
        var cAbw = Col("Abwicklungsmaß 90°", "Abwicklungsmaß 90", "Abwicklungsmass 90°", "Abwicklungsmass 90");
        var cMinS = Col("Schenkelmaß minimal", "Schenkelmas minimal");

        static decimal? ReadDecimal(IXLCell cell)
        {
            if (cell.IsEmpty()) return null;
            var s = cell.GetString().Trim();
            if (string.IsNullOrWhiteSpace(s))
            {
                if (cell.TryGetValue<double>(out var d)) return (decimal)d;
                return null;
            }

            s = s.Replace(",", ".");
            if (decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                return v;

            if (cell.TryGetValue<double>(out var dd)) return (decimal)dd;
            return null;
        }

        static string? ReadString(IXLCell cell)
        {
            if (cell.IsEmpty()) return null;
            var s = cell.GetString().Trim();
            return string.IsNullOrWhiteSpace(s) ? null : s;
        }

        var rows = new List<RuleRow>();
        foreach (var row in ws.RowsUsed().Skip(1))
        {
            var material = ReadString(row.Cell(cMaterial));
            if (material is null) continue;

            var t = ReadDecimal(row.Cell(cT)) ?? 0m;
            var v = ReadString(row.Cell(cV)) ?? "UNI";

            rows.Add(new RuleRow(
                Material: material,
                ThicknessMm: t,
                PrismaV: v,
                Zuschlagverfahren: ReadString(row.Cell(cZ)),
                Biegeradien: ReadString(row.Cell(cR)),
                Massabzug: ReadDecimal(row.Cell(cAbzug)),
                Sollmass90: ReadDecimal(row.Cell(cSoll)),
                AbwicklungsmassOf90: ReadDecimal(row.Cell(cAbw)),
                MinSchenkelMm: ReadDecimal(row.Cell(cMinS))
            ));
        }

        return rows;
    }

    public RuleRow? FindBestRule(IEnumerable<RuleRow> rules, string material, decimal thicknessMm, string? prismaV)
    {
        var mat = rules.Where(r => string.Equals(r.Material, material, StringComparison.OrdinalIgnoreCase)).ToList();
        if (mat.Count == 0) return null;

        const decimal tol = 0.06m;
        var exact = mat.Where(r => Math.Abs(r.ThicknessMm - thicknessMm) <= tol).ToList();
        var candidates = exact.Count > 0 ? exact : mat.OrderBy(r => Math.Abs(r.ThicknessMm - thicknessMm)).Take(10).ToList();

        if (!string.IsNullOrWhiteSpace(prismaV))
        {
            var pv = prismaV.Trim();
            var vmatch = candidates.FirstOrDefault(r => string.Equals(r.PrismaV.Trim(), pv, StringComparison.OrdinalIgnoreCase));
            if (vmatch is not null) return vmatch;
        }

        return candidates.FirstOrDefault();
    }

    public string? SuggestPrismaV(IEnumerable<RuleRow> rules, string material, decimal thicknessMm)
    {
        return FindBestRule(rules, material, thicknessMm, null)?.PrismaV;
    }
}

