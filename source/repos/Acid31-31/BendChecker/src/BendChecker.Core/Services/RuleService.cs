using BendChecker.Core.Models;
using ClosedXML.Excel;

namespace BendChecker.Core.Services;

public sealed class RuleService
{
    public List<RuleRow> LoadRules(string xlsxPath, string? sheetName = null)
    {
        using var wb = new XLWorkbook(xlsxPath);
        var ws = sheetName is null ? wb.Worksheets.First() : wb.Worksheet(sheetName);

        var headerRow = ws.FirstRowUsed();
        var headerMap = headerRow.CellsUsed()
            .ToDictionary(c => c.GetString().Trim(), c => c.Address.ColumnNumber, StringComparer.OrdinalIgnoreCase);

        int Col(string name)
        {
            if (!headerMap.TryGetValue(name, out var col))
                throw new InvalidOperationException($"Excel-Spalte fehlt: '{name}'. Vorhanden: {string.Join(", ", headerMap.Keys)}");
            return col;
        }

        var cMaterial = Col("Material");
        var cT = Col("Materialstärke");
        var cV = Col("Prisma V");
        var cZ = Col("Zuschlagverfahren");
        var cR = Col("Biegeradien");
        var cAbzug = Col("Maßabzug");
        var cSoll = Col("Sollmaß 90°");
        var cAbw = Col("Abwicklungsmaß 90°");
        var cMinS = Col("Schenkelmaß minimal");

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
                Abwicklungsmaß90: ReadDecimal(row.Cell(cAbw)),
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
}
