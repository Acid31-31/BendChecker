using System.Globalization;
using System.Text;
using Occt;

namespace BendChecker.Core.Services;

/// <summary>
/// Exports a STEP file to a HiCAD-style GEO 2D flat-pattern file.
/// The output file has the same base name and is written next to the input STEP.
/// Phase 1: planar / cylindrical face classification, 2D projection onto XY,
/// bend-zone extraction (angle, radius, thickness), convex-hull outer contour.
/// </summary>
public sealed class StepToGeoExporter
{
    // ── geometry / algorithm constants ───────────────────────────────────────

    /// <summary>Minimum dot-product (|cos θ|) for two normals to be considered parallel.</summary>
    private const double ParallelNormalThreshold = 0.999;

    /// <summary>Minimum dot-product for a face normal to be considered aligned with the dominant base normal.</summary>
    private const double DominantNormalAlignmentThreshold = 0.85;

    /// <summary>Maximum dot-product for two normals to be clustered into the same direction group.</summary>
    private const double NormalClusteringThreshold = 0.98;

    /// <summary>Maximum dot-product for two normals to be considered sufficiently different (for bend adjacent-plane selection).</summary>
    private const double DifferentNormalThreshold = 0.98;

    /// <summary>Maximum |cos θ| for a planar face normal to be considered perpendicular to a cylinder axis.</summary>
    private const double AxisPerpendicularityThreshold = 0.5;

    /// <summary>Spatial tolerance used when deduplicating projected 2-D points (mm).</summary>
    private const double PointDeduplicationTolerance = 0.1;

    /// <summary>Spatial tolerance used when registering bend-line endpoints as 2-D points (mm).</summary>
    private const double BendLinePointTolerance = 0.5;

    /// <summary>Minimum length for a cross-product vector to be considered non-degenerate.</summary>
    private const double CrossProductTolerance = 1e-12;

    // ── internal geometry records ────────────────────────────────────────────

    private sealed record PlaneData(gp_Dir Normal, gp_Pnt Origin, List<gp_Pnt> Vertices);

    private sealed record CylData(gp_Dir AxisDir, gp_Pnt AxisPoint, double Radius, List<gp_Pnt> Vertices);

    private sealed record Pt2d(int Id, double X, double Y);

    private sealed record GeoEdge(int P1, int P2);

    private sealed record BendInfo(double AngleDeg, double ThicknessMm, double RadiusMm, int LineP1, int LineP2);

    // ── public API ───────────────────────────────────────────────────────────

    /// <summary>
    /// Exports the STEP file at <paramref name="stepPath"/> to a GEO file in the same folder
    /// with the same base name (extension changed to .GEO).
    /// </summary>
    /// <returns>Full path of the generated GEO file.</returns>
    public string Export(string stepPath, CancellationToken ct = default)
    {
        if (!File.Exists(stepPath))
            throw new FileNotFoundException("STEP-Datei nicht gefunden.", stepPath);

        var partName = Path.GetFileNameWithoutExtension(stepPath);
        var geoPath = Path.Combine(
            Path.GetDirectoryName(stepPath) ?? ".",
            partName + ".GEO");

        // ── 1. Load STEP via XCAF ────────────────────────────────────────────
        ct.ThrowIfCancellationRequested();

        var reader = new STEPCAFControl_Reader
        {
            NameMode = true,
            ColorMode = true,
            LayerMode = true
        };

        var status = reader.ReadFile(stepPath);
        if (status != IFSelect_ReturnStatus.IFSelect_RetDone)
            throw new InvalidOperationException($"STEP konnte nicht gelesen werden ({status}).");

        var document = new TDocStd_Document(new TCollection_ExtendedString("BinXCAF"));
        if (!reader.Transfer(document))
            throw new InvalidOperationException("STEP-Transfer in XCAF-Dokument fehlgeschlagen.");

        var shapeTool = XCAFDoc_DocumentTool.ShapeTool(document.Main);
        shapeTool.GetFreeShapes(out var freeShapes);
        if (freeShapes.IsEmpty)
            throw new InvalidOperationException("STEP enthält keine freien Shapes.");

        var rootShape = XCAFDoc_ShapeTool.GetOneShape(freeShapes);
        if (rootShape is null || rootShape.IsNull)
            throw new InvalidOperationException("STEP enthält keine darstellbare Geometrie.");

        ct.ThrowIfCancellationRequested();

        // ── 2. Triangulate (needed for vertex collection) ────────────────────
        var bounds = rootShape.BoundingBox;
        var diagonal = bounds.CornerMin.Distance(bounds.CornerMax);
        var deflection = diagonal > 0 ? Math.Max(diagonal / 600d, 0.05d) : 0.1d;
        _ = new BRepMesh_IncrementalMesh(rootShape, deflection, false, 0.5, true);

        ct.ThrowIfCancellationRequested();

        // ── 3. Classify faces ────────────────────────────────────────────────
        var planarFaces = new List<PlaneData>();
        var cylindricalFaces = new List<CylData>();

        CollectFaceGeometry(rootShape, planarFaces, cylindricalFaces, ct);

        if (planarFaces.Count == 0)
            throw new InvalidOperationException("Keine planaren Flächen im STEP gefunden – kein Blechbauteil erkannt.");

        // ── 4. Estimate thickness ─────────────────────────────────────────────
        var thicknessMm = EstimateThickness(planarFaces);

        // ── 5. Choose dominant base normal → define local XY coordinate system ─
        var baseNormal = FindDominantNormal(planarFaces);
        var (xAxis2d, yAxis2d) = BuildXYAxes(baseNormal);

        // ── 6. Project all 3-D points from planar faces to local 2-D XY ─────
        var raw2d = new List<(double X, double Y)>();
        foreach (var pf in planarFaces)
        {
            // Only project faces whose normal is roughly aligned with the base normal
            // (top/bottom faces of the sheet).
            var dot = Math.Abs(pf.Normal.Dot(baseNormal));
            if (dot < DominantNormalAlignmentThreshold)
                continue;

            foreach (var v in pf.Vertices)
            {
                raw2d.Add(ProjectPoint(v, xAxis2d, yAxis2d));
            }
        }

        // If the dominant-plane filter gave nothing, include everything
        if (raw2d.Count < 3)
        {
            raw2d.Clear();
            foreach (var pf in planarFaces)
            {
                foreach (var v in pf.Vertices)
                    raw2d.Add(ProjectPoint(v, xAxis2d, yAxis2d));
            }
        }

        // ── 7. Deduplicate and build 2-D point list ──────────────────────────
        var uniquePts = BuildUniquePoints(raw2d, tolerance: PointDeduplicationTolerance);
        if (uniquePts.Count < 2)
            throw new InvalidOperationException("Zu wenige Punkte für die GEO-Ausgabe.");

        // ── 8. Outer contour: convex hull of projected points ────────────────
        var hullIdx = ConvexHull(uniquePts);
        var contourEdges = new List<GeoEdge>();
        for (var i = 0; i < hullIdx.Count; i++)
        {
            var a = hullIdx[i];
            var b = hullIdx[(i + 1) % hullIdx.Count];
            contourEdges.Add(new GeoEdge(uniquePts[a].Id, uniquePts[b].Id));
        }

        // ── 9. Extract bend zones from cylindrical faces ─────────────────────
        var bends = ExtractBendZones(cylindricalFaces, planarFaces, xAxis2d, yAxis2d, uniquePts, thicknessMm);

        // ── 10. Write GEO ─────────────────────────────────────────────────────
        var geoContent = WriteGeo(partName, uniquePts, hullIdx, contourEdges, bends, thicknessMm);
        File.WriteAllText(geoPath, geoContent, Encoding.ASCII);

        return geoPath;
    }

    // ── face geometry collection ─────────────────────────────────────────────

    private static void CollectFaceGeometry(
        TopoDS_Shape rootShape,
        List<PlaneData> planarFaces,
        List<CylData> cylindricalFaces,
        CancellationToken ct)
    {
        unsafe
        {
            var faceExplorer = new TopExp_Explorer(rootShape, TopAbs_ShapeEnum.TopAbs_FACE);
            while (faceExplorer.More)
            {
                ct.ThrowIfCancellationRequested();

                try
                {
                    var faceShape = faceExplorer.Current;
                    var face = TopoDS_Face.Cast((nint)faceShape.NativeInstance);
                    var adaptor = new BRepAdaptor_Surface(face);

                    if (adaptor.Type == GeomAbs_SurfaceType.GeomAbs_Plane)
                    {
                        var pln = adaptor.Plane;
                        var normal = NormalizeToUpperHemisphere(pln.Axis.Direction);
                        var vertices = CollectFaceVertices(face);
                        if (vertices.Count > 0)
                            planarFaces.Add(new PlaneData(normal, pln.Location, vertices));
                    }
                    else if (adaptor.Type == GeomAbs_SurfaceType.GeomAbs_Cylinder)
                    {
                        TryAddCylinderFace(adaptor, face, cylindricalFaces);
                    }
                }
                catch
                {
                    // skip problematic faces
                }

                faceExplorer.Next();
            }
        }
    }

    private static void TryAddCylinderFace(BRepAdaptor_Surface adaptor, TopoDS_Face face, List<CylData> cylindricalFaces)
    {
        try
        {
            var cylinder = adaptor.Cylinder;
            var axis = cylinder.Axis;
            var vertices = CollectFaceVertices(face);
            if (vertices.Count > 0)
            {
                cylindricalFaces.Add(new CylData(
                    axis.Direction,
                    axis.Location,
                    cylinder.Radius,
                    vertices));
            }
        }
        catch
        {
            // Cylinder property not accessible – estimate from triangulated vertices.
            // NOTE: The radius is estimated as the average XY-plane distance from the centroid.
            // This assumes the cylinder axis is approximately aligned with the world Z-axis;
            // for differently oriented cylinders the estimate may be inaccurate.
            // The axis direction is hard-coded to (0,0,1) as a safe fallback.
            try
            {
                var vertices = CollectFaceVertices(face);
                if (vertices.Count >= 3)
                {
                    var cx = vertices.Average(v => v.X);
                    var cy = vertices.Average(v => v.Y);
                    var cz = vertices.Average(v => v.Z);
                    var centroid = new gp_Pnt(cx, cy, cz);
                    var radiusEst = vertices.Average(v => Math.Sqrt((v.X - cx) * (v.X - cx) + (v.Y - cy) * (v.Y - cy)));
                    if (radiusEst > 0.01)
                    {
                        cylindricalFaces.Add(new CylData(
                            new gp_Dir(0, 0, 1), // fallback axis: world Z
                            centroid,
                            radiusEst,
                            vertices));
                    }
                }
            }
            catch
            {
                // ignore
            }
        }
    }

    private static List<gp_Pnt> CollectFaceVertices(TopoDS_Face face)
    {
        var pts = new List<gp_Pnt>();
        try
        {
            var triangulation = BRep_Tool.Triangulation(face, out var location);
            if (triangulation is null || triangulation.NbNodes == 0)
                return pts;

            var transform = location.Transformation;
            for (var i = 1; i <= triangulation.NbNodes; i++)
            {
                var node = triangulation.Node(i);
                pts.Add(node.Transformed(transform));
            }
        }
        catch
        {
            // ignore triangulation errors for this face
        }

        return pts;
    }

    // ── geometry utilities ────────────────────────────────────────────────────

    private static decimal? EstimateThickness(List<PlaneData> planarFaces)
    {
        // Find closest pair of parallel planes
        double? best = null;
        for (var i = 0; i < planarFaces.Count; i++)
        {
            for (var j = i + 1; j < planarFaces.Count; j++)
            {
                var nA = planarFaces[i].Normal;
                var nB = planarFaces[j].Normal;
                if (Math.Abs(nA.Dot(nB)) < ParallelNormalThreshold)
                    continue;

                var dist = planarFaces[i].Origin.Distance(planarFaces[j].Origin);
                if (dist < 0.01)
                    continue;

                // Use the normal direction to project the distance properly
                var vec = new gp_Vec(planarFaces[i].Origin, planarFaces[j].Origin);
                var d = Math.Abs(vec.Dot(new gp_Vec(nA.X, nA.Y, nA.Z)));
                if (d < 0.01)
                    continue;

                if (best is null || d < best.Value)
                    best = d;
            }
        }

        return best is null ? null : (decimal)best.Value;
    }

    private static gp_Dir FindDominantNormal(List<PlaneData> planarFaces)
    {
        // Cluster normals and pick the most frequent direction
        var clusters = new List<(gp_Dir Dir, int Count)>();

        foreach (var pf in planarFaces)
        {
            var n = pf.Normal;
            var found = false;
            for (var i = 0; i < clusters.Count; i++)
            {
                if (Math.Abs(clusters[i].Dir.Dot(n)) > NormalClusteringThreshold)
                {
                    clusters[i] = (clusters[i].Dir, clusters[i].Count + 1);
                    found = true;
                    break;
                }
            }

            if (!found)
                clusters.Add((n, 1));
        }

        if (clusters.Count == 0)
            return new gp_Dir(0, 0, 1);

        var best = clusters.OrderByDescending(c => c.Count).First();
        return best.Dir;
    }

    private static (gp_Dir XAxis, gp_Dir YAxis) BuildXYAxes(gp_Dir normal)
    {
        // Build an orthonormal basis perpendicular to 'normal' using manual cross-product.
        // gp_Dir is always a unit vector in OCCT, so no explicit normalization is needed
        // once we divide by the length.
        var xAxis = CrossDir(normal, Math.Abs(normal.X) < 0.9 ? new gp_Dir(1, 0, 0) : new gp_Dir(0, 1, 0));
        var yAxis = CrossDir(normal, xAxis);
        return (xAxis, yAxis);
    }

    /// <summary>
    /// Flips a face normal so that it points into the upper hemisphere (Z > 0, or if Z≈0
    /// then Y > 0). This ensures normals of opposite faces of the same sheet metal part get
    /// grouped together during clustering, avoiding double-counting the same plane direction.
    /// </summary>
    private static gp_Dir NormalizeToUpperHemisphere(gp_Dir d)
    {
        if (d.Z < 0 || (Math.Abs(d.Z) < 1e-6 && d.Y < 0))
            return d.Reversed;
        return d;
    }

    /// <summary>Computes a × b and returns the normalized result as a gp_Dir.</summary>
    private static gp_Dir CrossDir(gp_Dir a, gp_Dir b)
    {
        var x = a.Y * b.Z - a.Z * b.Y;
        var y = a.Z * b.X - a.X * b.Z;
        var z = a.X * b.Y - a.Y * b.X;
        var len = Math.Sqrt(x * x + y * y + z * z);
        if (len < CrossProductTolerance)
            return new gp_Dir(0, 0, 1);
        return new gp_Dir(x / len, y / len, z / len);
    }

    private static (double X, double Y) ProjectPoint(gp_Pnt pt, gp_Dir xAxis, gp_Dir yAxis)
    {
        var x = pt.X * xAxis.X + pt.Y * xAxis.Y + pt.Z * xAxis.Z;
        var y = pt.X * yAxis.X + pt.Y * yAxis.Y + pt.Z * yAxis.Z;
        return (x, y);
    }

    // ── 2D point management ───────────────────────────────────────────────────

    private static List<Pt2d> BuildUniquePoints(List<(double X, double Y)> raw, double tolerance)
    {
        var result = new List<Pt2d>();
        var nextId = 1;

        foreach (var (rx, ry) in raw)
        {
            var found = false;
            foreach (var existing in result)
            {
                var dx = existing.X - rx;
                var dy = existing.Y - ry;
                if (Math.Sqrt(dx * dx + dy * dy) < tolerance)
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                result.Add(new Pt2d(nextId++, rx, ry));
            }
        }

        return result;
    }

    private static int? FindNearestPoint(List<Pt2d> pts, double x, double y, double tolerance)
    {
        double bestDist = double.MaxValue;
        int? bestId = null;
        foreach (var p in pts)
        {
            var d = Math.Sqrt((p.X - x) * (p.X - x) + (p.Y - y) * (p.Y - y));
            if (d < tolerance && d < bestDist)
            {
                bestDist = d;
                bestId = p.Id;
            }
        }

        return bestId;
    }

    // ── convex hull (Andrew's monotone chain) ─────────────────────────────────

    private static List<int> ConvexHull(List<Pt2d> pts)
    {
        if (pts.Count <= 2)
            return Enumerable.Range(0, pts.Count).ToList();

        var indexed = pts
            .Select((p, i) => (p, i))
            .OrderBy(t => t.p.X)
            .ThenBy(t => t.p.Y)
            .ToList();

        static double Cross(Pt2d o, Pt2d a, Pt2d b) =>
            (a.X - o.X) * (b.Y - o.Y) - (a.Y - o.Y) * (b.X - o.X);

        var lower = new List<int>();
        foreach (var (p, idx) in indexed)
        {
            while (lower.Count >= 2 && Cross(pts[lower[^2]], pts[lower[^1]], p) <= 0)
                lower.RemoveAt(lower.Count - 1);
            lower.Add(idx);
        }

        var upper = new List<int>();
        for (var i = indexed.Count - 1; i >= 0; i--)
        {
            var (p, idx) = indexed[i];
            while (upper.Count >= 2 && Cross(pts[upper[^2]], pts[upper[^1]], p) <= 0)
                upper.RemoveAt(upper.Count - 1);
            upper.Add(idx);
        }

        lower.RemoveAt(lower.Count - 1);
        upper.RemoveAt(upper.Count - 1);
        lower.AddRange(upper);
        return lower;
    }

    // ── bend zone extraction ─────────────────────────────────────────────────

    private static List<BendInfo> ExtractBendZones(
        List<CylData> cylindricalFaces,
        List<PlaneData> planarFaces,
        gp_Dir xAxis,
        gp_Dir yAxis,
        List<Pt2d> uniquePts,
        decimal? thicknessMm)
    {
        var bends = new List<BendInfo>();
        var t = thicknessMm is null ? 1.5 : (double)thicknessMm;

        foreach (var cyl in cylindricalFaces)
        {
            try
            {
                // Find the two planar faces whose normals are most perpendicular to the cylinder axis
                // These are the two flat flanges connected by this bend.
                var adjacentPlanes = FindAdjacentPlanarFaces(cyl, planarFaces);
                if (adjacentPlanes.Count < 2)
                    continue;

                // Bend angle = angle between the two adjacent planar normals (supplement)
                var n1 = adjacentPlanes[0].Normal;
                var n2 = adjacentPlanes[1].Normal;
                var cosAngle = Math.Clamp(n1.Dot(n2), -1.0, 1.0);
                var angleBetweenNormals = Math.Acos(cosAngle) * 180.0 / Math.PI;
                // The bend angle (as measured in sheet metal) is 180° – angleBetweenNormals when normals point outward
                var bendAngleDeg = Math.Round(180.0 - angleBetweenNormals, 3);
                if (bendAngleDeg <= 0 || bendAngleDeg >= 180)
                    bendAngleDeg = 90.0; // Geometry could not determine a valid angle; defaulting to 90° (TODO: log a warning)

                // Project the cylinder axis endpoints to 2D as the bend line
                var (axPt1, axPt2) = GetCylinderAxisEndpoints(cyl);
                var (ax1x, ax1y) = ProjectPoint(axPt1, xAxis, yAxis);
                var (ax2x, ax2y) = ProjectPoint(axPt2, xAxis, yAxis);

                // Register the two axis endpoints as unique 2D points
                var id1 = RegisterPoint(uniquePts, ax1x, ax1y, BendLinePointTolerance);
                var id2 = RegisterPoint(uniquePts, ax2x, ax2y, BendLinePointTolerance);

                if (id1 == id2)
                    continue;

                bends.Add(new BendInfo(bendAngleDeg, t, cyl.Radius, id1, id2));
            }
            catch
            {
                // skip problematic bend zones
            }
        }

        return bends;
    }

    private static List<PlaneData> FindAdjacentPlanarFaces(CylData cyl, List<PlaneData> planarFaces)
    {
        // Adjacent flat faces are those whose centroid is closest to the cylinder axis
        // and whose normal is roughly perpendicular to the cylinder axis.
        var results = new List<(PlaneData Plane, double Score)>();

        foreach (var pf in planarFaces)
        {
            // Normal should be perpendicular to cylinder axis (dot near 0)
            var dotWithAxis = Math.Abs(pf.Normal.Dot(cyl.AxisDir));
            if (dotWithAxis > AxisPerpendicularityThreshold)
                continue; // this face normal is too aligned with the axis

            // Centroid of the planar face
            if (pf.Vertices.Count == 0)
                continue;

            var cx = pf.Vertices.Average(v => v.X);
            var cy = pf.Vertices.Average(v => v.Y);
            var cz = pf.Vertices.Average(v => v.Z);
            var centroid = new gp_Pnt(cx, cy, cz);

            // Distance from centroid to cylinder axis line
            var toCentroid = new gp_Vec(cyl.AxisPoint, centroid);
            var axVec = new gp_Vec(cyl.AxisDir.X, cyl.AxisDir.Y, cyl.AxisDir.Z);
            var proj = toCentroid.Dot(axVec);
            var closestOnAxis = new gp_Pnt(
                cyl.AxisPoint.X + cyl.AxisDir.X * proj,
                cyl.AxisPoint.Y + cyl.AxisDir.Y * proj,
                cyl.AxisPoint.Z + cyl.AxisDir.Z * proj);
            var distToAxis = centroid.Distance(closestOnAxis);

            // Good candidates are close to the cylinder's radius distance from axis
            var radialDiff = Math.Abs(distToAxis - cyl.Radius);
            results.Add((pf, radialDiff));
        }

        // Take the two best candidates with different normals
        var sorted = results.OrderBy(r => r.Score).Select(r => r.Plane).ToList();
        var selected = new List<PlaneData>();
        foreach (var pf in sorted)
        {
            if (selected.Count == 0)
            {
                selected.Add(pf);
            }
            else if (selected.Count == 1)
            {
                // Check that the normal is different enough
                if (Math.Abs(selected[0].Normal.Dot(pf.Normal)) < DifferentNormalThreshold)
                    selected.Add(pf);
            }
            else
            {
                break;
            }
        }

        return selected;
    }

    private static (gp_Pnt P1, gp_Pnt P2) GetCylinderAxisEndpoints(CylData cyl)
    {
        if (cyl.Vertices.Count == 0)
        {
            // Fallback: dummy axis endpoints
            return (cyl.AxisPoint, new gp_Pnt(
                cyl.AxisPoint.X + cyl.AxisDir.X * 10,
                cyl.AxisPoint.Y + cyl.AxisDir.Y * 10,
                cyl.AxisPoint.Z + cyl.AxisDir.Z * 10));
        }

        // Project all face vertices onto the cylinder axis to get min/max extent
        double minT = double.MaxValue, maxT = double.MinValue;
        foreach (var v in cyl.Vertices)
        {
            var toV = new gp_Vec(cyl.AxisPoint, v);
            var axVec = new gp_Vec(cyl.AxisDir.X, cyl.AxisDir.Y, cyl.AxisDir.Z);
            var t = toV.Dot(axVec);
            if (t < minT) minT = t;
            if (t > maxT) maxT = t;
        }

        var p1 = new gp_Pnt(
            cyl.AxisPoint.X + cyl.AxisDir.X * minT,
            cyl.AxisPoint.Y + cyl.AxisDir.Y * minT,
            cyl.AxisPoint.Z + cyl.AxisDir.Z * minT);
        var p2 = new gp_Pnt(
            cyl.AxisPoint.X + cyl.AxisDir.X * maxT,
            cyl.AxisPoint.Y + cyl.AxisDir.Y * maxT,
            cyl.AxisPoint.Z + cyl.AxisDir.Z * maxT);
        return (p1, p2);
    }

    private static int RegisterPoint(List<Pt2d> pts, double x, double y, double tolerance)
    {
        var nearest = FindNearestPoint(pts, x, y, tolerance);
        if (nearest is not null)
            return nearest.Value;

        var newId = pts.Count == 0 ? 1 : pts.Max(p => p.Id) + 1;
        pts.Add(new Pt2d(newId, x, y));
        return newId;
    }

    // ── GEO writer ────────────────────────────────────────────────────────────

    private static string WriteGeo(
        string partName,
        List<Pt2d> points,
        List<int> hullIndices,
        List<GeoEdge> contourEdges,
        List<BendInfo> bends,
        decimal? thicknessMm)
    {
        var sb = new StringBuilder();
        var ic = CultureInfo.InvariantCulture;
        var date = DateTime.Now.ToString("yyyy-MM-dd", ic);

        // Bounding box of 2D points
        double minX = 0, minY = 0, maxX = 0, maxY = 0;
        if (points.Count > 0)
        {
            minX = points.Min(p => p.X);
            minY = points.Min(p => p.Y);
            maxX = points.Max(p => p.X);
            maxY = points.Max(p => p.Y);
        }

        var width = maxX - minX;
        var height = maxY - minY;

        // Center points at origin
        var offsetX = minX + width / 2.0;
        var offsetY = minY + height / 2.0;

        // ── #~1 Header ───────────────────────────────────────────────────────
        sb.AppendLine("#~1");
        sb.AppendLine($"DATUM = {date}");
        sb.AppendLine($"0..{width.ToString("0.###", ic)}");
        sb.AppendLine($"0..{height.ToString("0.###", ic)}");

        // ── #~3 Metadata ─────────────────────────────────────────────────────
        sb.AppendLine("#~3");
        sb.AppendLine(partName);
        sb.AppendLine("LASER");

        // ── #~31 Points ──────────────────────────────────────────────────────
        sb.AppendLine("#~31");
        foreach (var pt in points)
        {
            // Center coordinates around origin (HiCAD convention)
            var cx = pt.X - offsetX;
            var cy = pt.Y - offsetY;
            sb.AppendLine("P");
            sb.AppendLine(pt.Id.ToString(ic));
            sb.AppendLine($"{cx.ToString("0.000000000", ic)} {cy.ToString("0.000000000", ic)} 0.000000000");
        }

        // ── #~331 Contour segments ───────────────────────────────────────────
        sb.AppendLine("#~331");
        var segIdx = 1;
        foreach (var edge in contourEdges)
        {
            sb.AppendLine("LIN");
            sb.AppendLine(segIdx.ToString(ic));
            sb.AppendLine($"{edge.P1} {edge.P2}");
            segIdx++;
        }

        // ── #~37 Bend zones ──────────────────────────────────────────────────
        foreach (var bend in bends)
        {
            sb.AppendLine("#~37");
            sb.AppendLine(bend.AngleDeg.ToString("0.000000000", ic));
            sb.AppendLine((thicknessMm is null ? bend.ThicknessMm : (double)thicknessMm).ToString("0.000000000", ic));
            sb.AppendLine(bend.RadiusMm.ToString("0.000000000", ic));
            // Bend line as a LIN segment
            sb.AppendLine("LIN");
            sb.AppendLine(segIdx.ToString(ic));
            sb.AppendLine($"{bend.LineP1} {bend.LineP2}");
            segIdx++;
            sb.AppendLine("#~BIEG_END");
        }

        return sb.ToString();
    }
}
