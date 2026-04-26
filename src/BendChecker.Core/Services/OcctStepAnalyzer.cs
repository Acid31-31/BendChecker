using System.Globalization;
using System.Text.RegularExpressions;
using BendChecker.Core.Models;
using Occt;

namespace BendChecker.Core.Services;

public sealed class OcctStepAnalyzer : IStepAnalyzer, IStepPreviewLoader
{
    public Task<bool> CanOpenAsync(string stepPath, CancellationToken ct)
    {
        return Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();

            if (!IsSupportedStepFile(stepPath))
                return false;

            try
            {
                _ = ReadRootShape(stepPath, ct);
                return true;
            }
            catch
            {
                return false;
            }
        }, ct);
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

    public Task<StepPreviewScene> LoadPreviewAsync(string stepPath, CancellationToken ct)
    {
        return Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();

            var shape = ReadRootShape(stepPath, ct);
            var triangleVertices = Triangulate(shape, ct);
            if (triangleVertices.Count == 0)
                throw new InvalidOperationException("In der STEP-Datei wurde keine darstellbare 3D-Geometrie gefunden.");

            return new StepPreviewScene(triangleVertices.ToArray());
        }, ct);
    }

    private static bool IsSupportedStepFile(string stepPath)
    {
        return File.Exists(stepPath) &&
               (stepPath.EndsWith(".step", StringComparison.OrdinalIgnoreCase) ||
                stepPath.EndsWith(".stp", StringComparison.OrdinalIgnoreCase));
    }

    private static TopoDS_Shape ReadRootShape(string stepPath, CancellationToken ct)
    {
        if (!IsSupportedStepFile(stepPath))
            throw new FileNotFoundException("STEP-Datei wurde nicht gefunden oder hat keine gueltige Endung.", stepPath);

        ct.ThrowIfCancellationRequested();

        var reader = new STEPCAFControl_Reader
        {
            NameMode = true,
            ColorMode = true,
            LayerMode = true,
            MatMode = true
        };

        var status = reader.ReadFile(stepPath);
        if (status != IFSelect_ReturnStatus.IFSelect_RetDone)
            throw new InvalidOperationException($"STEP-Datei konnte nicht gelesen werden ({status}).");

        var document = new TDocStd_Document(new TCollection_ExtendedString("BinXCAF"));
        if (!reader.Transfer(document))
            throw new InvalidOperationException("STEP-Datei konnte nicht in ein XCAF-Dokument uebertragen werden.");

        var shapeTool = XCAFDoc_DocumentTool.ShapeTool(document.Main);
        shapeTool.GetFreeShapes(out var freeShapes);
        if (freeShapes.IsEmpty)
            throw new InvalidOperationException("Die STEP-Datei enthaelt keine freien Shapes.");

        var shape = XCAFDoc_ShapeTool.GetOneShape(freeShapes);
        if (shape is null || shape.IsNull)
            throw new InvalidOperationException("Die STEP-Datei enthaelt keine darstellbare Geometrie.");

        return shape;
    }

    private static List<double> Triangulate(TopoDS_Shape shape, CancellationToken ct)
    {
        var bounds = shape.BoundingBox;
        var diagonal = bounds.CornerMin.Distance(bounds.CornerMax);
        var deflection = diagonal > 0 ? Math.Max(diagonal / 800d, 0.05d) : 0.1d;

        _ = new BRepMesh_IncrementalMesh(shape, deflection, false, 0.5, true);

        var triangleVertices = new List<double>();

        unsafe
        {
            var explorer = new TopExp_Explorer(shape, TopAbs_ShapeEnum.TopAbs_FACE);
            while (explorer.More)
            {
                ct.ThrowIfCancellationRequested();

                var faceShape = explorer.Current;
                var face = TopoDS_Face.Cast((nint)faceShape.NativeInstance);
                var triangulation = BRep_Tool.Triangulation(face, out var location);

                if (triangulation is not null && triangulation.NbTriangles > 0)
                {
                    var transform = location.Transformation;
                    var reversed = face.Orientation == TopAbs_Orientation.TopAbs_REVERSED;

                    for (var i = 1; i <= triangulation.NbTriangles; i++)
                    {
                        var triangle = triangulation.Triangle(i);
                        triangle.Get(out var n1, out var n2, out var n3);

                        if (reversed)
                            (n2, n3) = (n3, n2);

                        AppendPoint(triangleVertices, triangulation.Node(n1), transform);
                        AppendPoint(triangleVertices, triangulation.Node(n2), transform);
                        AppendPoint(triangleVertices, triangulation.Node(n3), transform);
                    }
                }

                explorer.Next();
            }
        }

        return triangleVertices;
    }

    private static void AppendPoint(List<double> target, gp_Pnt point, gp_Trsf transform)
    {
        var transformed = point.Transformed(transform);
        target.Add(transformed.X);
        target.Add(transformed.Y);
        target.Add(transformed.Z);
    }
}
