using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

namespace BendChecker.App;

public partial class GeoPreviewWindow : Window
{
    public GeoPreviewWindow(string geoText)
    {
        InitializeComponent();
        RawGeoText.Text = geoText;
        DrawGeo(geoText);
    }

    private void DrawGeo(string geoText)
    {
        try
        {
            var lines = geoText
                .Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();

            var pointsByName = new Dictionary<string, Point>(StringComparer.OrdinalIgnoreCase);
            var pointsById = new Dictionary<int, Point>();
            var loopsByNames = new List<List<string>>();
            var entities = new List<EdgeEntity>();

            for (var i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                var parts = line.Split([' ', '\t', ';'], StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 0)
                    continue;

                // Standard-Punktformat: "P12 x y"
                if (parts[0].StartsWith("P", StringComparison.OrdinalIgnoreCase) && parts.Length >= 3)
                {
                    if (TryParseGeoDouble(parts[1], out var sx) && TryParseGeoDouble(parts[2], out var sy))
                    {
                        pointsByName[parts[0]] = new Point(sx, sy);
                    }
                    continue;
                }

                // Dein GEO-Format: "P" + nächste Zeile ID + nächste Zeile X Y Z
                if (line.Equals("P", StringComparison.OrdinalIgnoreCase) && i + 2 < lines.Count)
                {
                    if (int.TryParse(lines[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var pid))
                    {
                        var nums = ExtractDoubles(lines[i + 2]);
                        if (nums.Count >= 2)
                        {
                            var p = new Point(nums[0], nums[1]);
                            pointsById[pid] = p;
                            pointsByName[$"P{pid}"] = p;
                            i += 2;
                            continue;
                        }
                    }
                }

                if ((parts[0].StartsWith("LOOP", StringComparison.OrdinalIgnoreCase) || parts[0].Equals("POLY", StringComparison.OrdinalIgnoreCase) || parts[0].Equals("L", StringComparison.OrdinalIgnoreCase)) && parts.Length >= 3)
                {
                    loopsByNames.Add(parts.Skip(1).ToList());
                    continue;
                }

                // Dein GEO-Entity-Format #~331: LIN/ARC + Header + IDs
                if (line.Equals("LIN", StringComparison.OrdinalIgnoreCase) && i + 2 < lines.Count)
                {
                    var nums = ExtractInts(lines[i + 2]);
                    if (nums.Count >= 2)
                    {
                        entities.Add(new EdgeEntity(nums[0], nums[1], null));
                        i += 2;
                    }
                    continue;
                }

                if (line.Equals("ARC", StringComparison.OrdinalIgnoreCase) && i + 2 < lines.Count)
                {
                    var nums = ExtractInts(lines[i + 2]);
                    if (nums.Count >= 3)
                    {
                        // GEO-Dialekt (#~331): [Support/OnArc, Start, End]
                        entities.Add(new EdgeEntity(nums[1], nums[2], nums[0]));
                        i += 2;
                    }
                    continue;
                }
            }

            var loopPoints = new List<List<Point>>();

            if (loopsByNames.Count > 0 && pointsByName.Count > 0)
            {
                foreach (var loop in loopsByNames)
                {
                    var poly = loop.Where(pointsByName.ContainsKey).Select(k => pointsByName[k]).ToList();
                    if (poly.Count >= 3)
                        loopPoints.Add(poly);
                }
            }

            // Fallback für #~331 + #~31
            if (loopPoints.Count == 0 && entities.Count > 0 && pointsById.Count > 0)
            {
                loopPoints = BuildLoopsFromEntities(entities, pointsById);
            }

            // Fallback: keine Loops, aber Punkte vorhanden
            if (loopPoints.Count == 0 && pointsByName.Count >= 3)
            {
                var ordered = pointsByName.Keys
                    .OrderBy(ExtractPointNumber)
                    .ThenBy(k => k, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var poly = ordered.Select(k => pointsByName[k]).ToList();
                if (poly.Count >= 3)
                    loopPoints.Add(poly);
            }

            if (loopPoints.Count == 0)
            {
                InfoText.Text = "GEO enthält keine zeichnungsfähigen Konturen.";
                return;
            }

            var allPts = loopPoints.SelectMany(p => p).ToList();
            var minX = allPts.Min(p => p.X);
            var minY = allPts.Min(p => p.Y);
            var maxX = allPts.Max(p => p.X);
            var maxY = allPts.Max(p => p.Y);
            var w = Math.Max(1.0, maxX - minX);
            var h = Math.Max(1.0, maxY - minY);

            const double canvasW = 900d;
            const double canvasH = 600d;
            const double margin = 30d;
            var scale = Math.Min((canvasW - 2 * margin) / w, (canvasH - 2 * margin) / h);

            var geometry = new PathGeometry { FillRule = FillRule.EvenOdd };

            foreach (var poly in loopPoints)
            {
                var figPts = poly.Select(p => new Point(
                    margin + (p.X - minX) * scale,
                    canvasH - margin - (p.Y - minY) * scale)).ToList();

                if (figPts.Count < 3)
                    continue;

                var fig = new PathFigure { StartPoint = figPts[0], IsClosed = true, IsFilled = true };
                for (var j = 1; j < figPts.Count; j++)
                    fig.Segments.Add(new LineSegment(figPts[j], true));
                geometry.Figures.Add(fig);
            }

            var path = new Path
            {
                Data = geometry,
                Fill = new SolidColorBrush(Color.FromArgb(90, 70, 120, 190)),
                Stroke = new SolidColorBrush(Color.FromRgb(30, 60, 100)),
                StrokeThickness = 2
            };

            PreviewCanvas.Children.Clear();
            PreviewCanvas.Children.Add(path);
            InfoText.Text = $"GEO Vorschau: {loopPoints.Count} Loop(s), Größe {w:0.##} x {h:0.##} mm";
        }
        catch (Exception ex)
        {
            InfoText.Text = "GEO Vorschaufehler: " + ex.Message;
        }
    }

    private static List<List<Point>> BuildLoopsFromEntities(IReadOnlyList<EdgeEntity> entities, IReadOnlyDictionary<int, Point> points)
    {
        const double tol = 1e-6;
        var loops = new List<List<Point>>();

        var segments = new List<List<Point>>();
        foreach (var e in entities)
        {
            if (!points.TryGetValue(e.StartId, out var ps) || !points.TryGetValue(e.EndId, out var pe))
                continue;

            // Für dieses GEO-Format keine Stütz-/Hilfspunkte in die Kontur übernehmen,
            // sonst entstehen Ausreißer (z. B. diagonale Spitzen an den Enden).
            var seg = new List<Point> { ps, pe };

            if (Distance(seg[0], seg[^1]) > tol)
                segments.Add(seg);
        }

        var used = new bool[segments.Count];

        List<Point> ReverseSegment(List<Point> s)
        {
            var copy = s.ToList();
            copy.Reverse();
            return copy;
        }

        for (var i = 0; i < segments.Count; i++)
        {
            if (used[i])
                continue;

            var path = new List<Point>(segments[i]);
            used[i] = true;

            var guard = 0;
            while (guard++ < segments.Count * 3)
            {
                var end = path[^1];
                var start = path[0];

                if (path.Count >= 4 && Distance(end, start) < tol)
                    break;

                var found = false;
                for (var j = 0; j < segments.Count; j++)
                {
                    if (used[j])
                        continue;

                    var seg = segments[j];
                    var s0 = seg[0];
                    var s1 = seg[^1];

                    if (Distance(end, s0) < tol)
                    {
                        for (var k = 1; k < seg.Count; k++)
                            path.Add(seg[k]);
                        used[j] = true;
                        found = true;
                        break;
                    }

                    if (Distance(end, s1) < tol)
                    {
                        var rev = ReverseSegment(seg);
                        for (var k = 1; k < rev.Count; k++)
                            path.Add(rev[k]);
                        used[j] = true;
                        found = true;
                        break;
                    }
                }

                if (!found)
                    break;
            }

            if (path.Count < 4 || Distance(path[0], path[^1]) >= tol)
                continue;

            path.RemoveAt(path.Count - 1); // closing duplicate

            // remove immediate duplicates
            var clean = new List<Point> { path[0] };
            for (var p = 1; p < path.Count; p++)
            {
                if (Distance(clean[^1], path[p]) >= tol)
                    clean.Add(path[p]);
            }

            if (clean.Count < 3)
                continue;

            var minX = clean.Min(p => p.X);
            var maxX = clean.Max(p => p.X);
            var minY = clean.Min(p => p.Y);
            var maxY = clean.Max(p => p.Y);
            if ((maxX - minX) < 0.1 || (maxY - minY) < 0.1)
                continue;

            loops.Add(clean);
        }

        return loops;
    }

    private readonly record struct EdgeEntity(int StartId, int EndId, int? MidId);

    private static double Distance(Point a, Point b)
    {
        var dx = a.X - b.X;
        var dy = a.Y - b.Y;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    private static List<int> ExtractInts(string text)
    {
        var matches = Regex.Matches(text, @"[-+]?\d+");
        var values = new List<int>();
        foreach (Match m in matches)
        {
            if (int.TryParse(m.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                values.Add(v);
        }

        return values;
    }

    private static List<double> ExtractDoubles(string text)
    {
        var matches = Regex.Matches(text, @"[-+]?\d+(?:[\.,]\d+)?(?:[eE][-+]?\d+)?");
        var values = new List<double>();
        foreach (Match m in matches)
        {
            if (TryParseGeoDouble(m.Value, out var d))
                values.Add(d);
        }

        return values;
    }

    private static bool TryParseGeoDouble(string raw, out double value)
    {
        raw = raw.Trim();

        if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            return true;

        if (double.TryParse(raw, NumberStyles.Any, CultureInfo.GetCultureInfo("de-DE"), out value))
            return true;

        var normalized = raw.Replace(',', '.');
        return double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
    }

    private static int ExtractPointNumber(string key)
    {
        var m = Regex.Match(key, "\\d+");
        return m.Success && int.TryParse(m.Value, out var n) ? n : int.MaxValue;
    }
}
