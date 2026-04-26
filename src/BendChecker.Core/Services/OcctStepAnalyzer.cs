using System.Globalization;
using System.Text.RegularExpressions;
using BendChecker.Core.Models;
using Occt;

namespace BendChecker.Core.Services;

public sealed class OcctStepAnalyzer : IStepAnalyzer
{
    private const byte DefaultRed = 185;
    private const byte DefaultGreen = 190;
    private const byte DefaultBlue = 205;
    private const byte DefaultAlpha = 255;

    public Task<bool> CanOpenAsync(string stepPath, CancellationToken ct)
    {
        return Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();

            if (!IsSupportedStepFile(stepPath))
                return false;

            try
            {
                _ = ReadDocumentContext(stepPath, ct);
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

            var thicknessFromName = StepAnalyzerStub.TryParseThicknessFromName(stepPath);
            if (thicknessFromName is not null)
                return thicknessFromName;

            var context = ReadDocumentContext(stepPath, ct);
            var thickness = TryEstimateThicknessMm(context.RootShape, ct);
            return thickness is null ? null : Math.Round(thickness.Value, 2, MidpointRounding.AwayFromZero);
        }, ct);
    }

    public Task<StepScene?> TryLoadSceneAsync(string stepPath, CancellationToken ct)
    {
        return Task.Run<StepScene?>(() =>
        {
            try
            {
                ct.ThrowIfCancellationRequested();

                var context = ReadDocumentContext(stepPath, ct);
                var parts = Triangulate(context.RootShape, context.ColorTool, ct);
                if (parts.Count == 0)
                    return null;

                return new StepScene(parts);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("STEP Preview failed: " + ex.Message, ex);
            }
        }, ct);
    }

    private static bool IsSupportedStepFile(string stepPath)
    {
        return File.Exists(stepPath) &&
               (stepPath.EndsWith(".step", StringComparison.OrdinalIgnoreCase) ||
                stepPath.EndsWith(".stp", StringComparison.OrdinalIgnoreCase));
    }

    private static StepDocumentContext ReadDocumentContext(string stepPath, CancellationToken ct)
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

        var colorTool = XCAFDoc_DocumentTool.ColorTool(document.Main);
        return new StepDocumentContext(shape, colorTool);
    }

    private static List<StepMeshPart> Triangulate(TopoDS_Shape shape, XCAFDoc_ColorTool? colorTool, CancellationToken ct)
    {
        var bounds = shape.BoundingBox;
        var diagonal = bounds.CornerMin.Distance(bounds.CornerMax);
        var deflection = diagonal > 0 ? Math.Max(diagonal / 800d, 0.05d) : 0.1d;

        _ = new BRepMesh_IncrementalMesh(shape, deflection, false, 0.5, true);

        var parts = new List<StepMeshPart>();

        unsafe
        {
            var explorer = new TopExp_Explorer(shape, TopAbs_ShapeEnum.TopAbs_FACE);
            while (explorer.More)
            {
                ct.ThrowIfCancellationRequested();

                try
                {
                    var faceShape = explorer.Current;
                    var face = TopoDS_Face.Cast((nint)faceShape.NativeInstance);
                    var triangulation = BRep_Tool.Triangulation(face, out var location);

                    if (triangulation is not null && triangulation.NbTriangles > 0)
                    {
                        var positions = new List<double>(triangulation.NbTriangles * 9);
                        var normals = new List<double>(triangulation.NbTriangles * 9);
                        var indices = new List<int>(triangulation.NbTriangles * 3);
                        var transform = location.Transformation;
                        var reversed = face.Orientation == TopAbs_Orientation.TopAbs_REVERSED;

                        for (var i = 1; i <= triangulation.NbTriangles; i++)
                        {
                            var triangle = triangulation.Triangle(i);
                            triangle.Get(out var n1, out var n2, out var n3);

                            if (reversed)
                                (n2, n3) = (n3, n2);

                            var p1 = TransformPoint(triangulation.Node(n1), transform);
                            var p2 = TransformPoint(triangulation.Node(n2), transform);
                            var p3 = TransformPoint(triangulation.Node(n3), transform);
                            var normal = CalculateNormal(p1, p2, p3);

                            AppendVertex(positions, normals, p1, normal, indices);
                            AppendVertex(positions, normals, p2, normal, indices);
                            AppendVertex(positions, normals, p3, normal, indices);
                        }

                        if (positions.Count > 0)
                        {
                            parts.Add(new StepMeshPart(
                                positions.ToArray(),
                                normals.ToArray(),
                                indices.ToArray(),
                                DefaultRed,
                                DefaultGreen,
                                DefaultBlue,
                                DefaultAlpha));
                        }
                    }
                }
                catch (Exception ex)
                {
                    _ = ex;
                    explorer.Next();
                    continue;
                }

                explorer.Next();
            }
        }

        return parts;
    }

    private static decimal? TryEstimateThicknessMm(TopoDS_Shape shape, CancellationToken ct)
    {
        var planes = new List<gp_Pln>();

        unsafe
        {
            var explorer = new TopExp_Explorer(shape, TopAbs_ShapeEnum.TopAbs_FACE);
            while (explorer.More)
            {
                ct.ThrowIfCancellationRequested();

                var faceShape = explorer.Current;
                var face = TopoDS_Face.Cast((nint)faceShape.NativeInstance);
                var adaptor = new BRepAdaptor_Surface(face);
                if (adaptor.Type == GeomAbs_SurfaceType.GeomAbs_Plane)
                    planes.Add(adaptor.Plane);

                explorer.Next();
            }
        }

        double? best = null;
        for (var i = 0; i < planes.Count; i++)
        {
            var normalA = planes[i].Axis.Direction;
            for (var j = i + 1; j < planes.Count; j++)
            {
                var normalB = planes[j].Axis.Direction;
                var alignment = Math.Abs(normalA.Dot(normalB));
                if (alignment < 0.999)
                    continue;

                var distance = planes[i].Distance(planes[j]);
                if (distance <= 0.01)
                    continue;

                if (best is null || distance < best.Value)
                    best = distance;
            }
        }

        if (best is not null)
            return (decimal)best.Value;

        var bounds = shape.BoundingBox;
        var cornerMin = bounds.CornerMin;
        var cornerMax = bounds.CornerMax;
        var fallback = new[]
            {
                Math.Abs(cornerMax.X - cornerMin.X),
                Math.Abs(cornerMax.Y - cornerMin.Y),
                Math.Abs(cornerMax.Z - cornerMin.Z)
            }
            .Where(v => v > 0.01)
            .DefaultIfEmpty(0)
            .Min();

        return fallback > 0 ? (decimal)fallback : null;
    }

    private static (byte Red, byte Green, byte Blue, byte Alpha) ResolveColor(XCAFDoc_ColorTool? colorTool, TopoDS_Shape faceShape, TopoDS_Shape rootShape)
    {
        if (colorTool is not null)
        {
            if (TryGetColor(colorTool, faceShape, out var faceColor))
                return faceColor;

            if (TryGetColor(colorTool, rootShape, out var shapeColor))
                return shapeColor;
        }

        return (DefaultRed, DefaultGreen, DefaultBlue, DefaultAlpha);
    }

    private static bool TryGetColor(XCAFDoc_ColorTool colorTool, TopoDS_Shape shape, out (byte Red, byte Green, byte Blue, byte Alpha) color)
    {
        if (colorTool.GetColor(shape, XCAFDoc_ColorType.XCAFDoc_ColorSurf, out Quantity_Color surfaceColor) ||
            colorTool.GetColor(shape, XCAFDoc_ColorType.XCAFDoc_ColorGen, out surfaceColor))
        {
            color = (
                ToByte(surfaceColor.Red),
                ToByte(surfaceColor.Green),
                ToByte(surfaceColor.Blue),
                DefaultAlpha);
            return true;
        }

        color = default;
        return false;
    }

    private static gp_Pnt TransformPoint(gp_Pnt point, gp_Trsf transform)
    {
        return point.Transformed(transform);
    }

    private static gp_Dir CalculateNormal(gp_Pnt p1, gp_Pnt p2, gp_Pnt p3)
    {
        var ux = p2.X - p1.X;
        var uy = p2.Y - p1.Y;
        var uz = p2.Z - p1.Z;
        var vx = p3.X - p1.X;
        var vy = p3.Y - p1.Y;
        var vz = p3.Z - p1.Z;

        var nx = uy * vz - uz * vy;
        var ny = uz * vx - ux * vz;
        var nz = ux * vy - uy * vx;
        var length = Math.Sqrt(nx * nx + ny * ny + nz * nz);
        if (length <= double.Epsilon)
            return new gp_Dir(0, 0, 1);

        return new gp_Dir(nx / length, ny / length, nz / length);
    }

    private static void AppendVertex(List<double> positions, List<double> normals, gp_Pnt point, gp_Dir normal, List<int> indices)
    {
        positions.Add(point.X);
        positions.Add(point.Y);
        positions.Add(point.Z);

        normals.Add(normal.X);
        normals.Add(normal.Y);
        normals.Add(normal.Z);

        indices.Add((positions.Count / 3) - 1);
    }

    private static byte ToByte(double value)
    {
        var scaled = (int)Math.Round(Math.Clamp(value, 0d, 1d) * 255d, MidpointRounding.AwayFromZero);
        return (byte)scaled;
    }

    private sealed record StepDocumentContext(TopoDS_Shape RootShape, XCAFDoc_ColorTool? ColorTool);
}
