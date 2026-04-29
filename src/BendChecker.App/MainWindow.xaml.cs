using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using BendChecker.Core.Models;
using BendChecker.Core.Services;
using HelixToolkit.Wpf;
using Microsoft.Win32;

namespace BendChecker.App;

public partial class MainWindow : Window
{
    private const int MaxReportLines = 120;
    private const int MaxReportEntryLength = 1200;
    private const bool PreviewSafeMode = true;
    private const bool DisableStepMeshRendering = false;

    private readonly RuleService _ruleService;
    private readonly IStepAnalyzer _stepAnalyzer;
    private readonly BendCheckService _svc;
    private readonly bool _isDesignMode;
    private readonly List<string> _reportLines = [];
    private readonly string _savedRulesPathFile = Path.Combine(App.GetReportFolderPath(), "rules-path.txt");
    private readonly string _rulesDbFile = Path.Combine(App.GetReportFolderPath(), "rules-db.json");
    private List<RuleRow> _importedRules = [];
    private HelixViewport3D? _stepViewport;
    private StepProbeResult? _lastProbeResult;
    private StepScene? _lastSceneForFlatPattern;
    private List<Point> _lastFlatPatternPolyline = [];
    private string _lastFlatPatternGeo = string.Empty;
    private bool _hasImportedRealGeo;
    private bool _hasGeneratedGeo;

    public MainWindow()
    {
        InitializeComponent();

        _isDesignMode = DesignerProperties.GetIsInDesignMode(this);
        _ruleService = new RuleService();
        _stepAnalyzer = new StepAnalyzerStub();
        _svc = new BendCheckService(_ruleService, new StepAnalyzerStub());

        if (_isDesignMode)
        {
            StatusText.Text = "Designer-Vorschau aktiv.";
            return;
        }

        AppendVisualReport("App gestartet.");
        AppendVisualReport(App.GetStartupDiagnostics());
        AppendVisualReport($"Report-Ordner: {App.GetReportFolderPath()}");
        ResetViewport();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        if (_isDesignMode)
        {
            StatusText.Text = "Designer-Vorschau aktiv.";
            return;
        }

        ShowStartupPreview();
        LoadRulesDb();
        TryRestoreRulesPath();
        StatusText.Text = "3D-Vorschau bereit.";
        AppendVisualReport("3D-Vorschau bereit.");
    }

    private async void PickStep_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "STEP (*.step;*.stp)|*.step;*.stp|All files (*.*)|*.*" };
        if (dlg.ShowDialog() != true)
            return;

        StepPathText.Text = dlg.FileName;
        AppendVisualReport($"STEP gewählt: {dlg.FileName}");
        App.MarkOperation("PickStep/LoadPreview");

        try
        {
            await LoadStepIntoUiAsync(dlg.FileName, CancellationToken.None);
            App.MarkIdle();
        }
        catch (Exception ex)
        {
            ResetViewport();
            ShowStartupPreview();
            StatusText.Text = "Fehler beim STEP lesen";
            AppendVisualReport($"STEP-Fehler: {ex}");
            App.MarkOperation("Fault: STEP lesen");
            MessageBox.Show(ex.ToString(), "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void PickRules_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*" };
        if (dlg.ShowDialog() != true)
            return;

        RulesPathText.Text = dlg.FileName;
        SaveRulesPath(dlg.FileName);

        try
        {
            _importedRules = _ruleService.LoadRulesFromAllSheets(dlg.FileName);
            SaveRulesDb();
            RulesInfoText.Text = $"Regeln importiert: {_importedRules.Count} Zeilen";
            AppendVisualReport($"Regeldatei importiert (alle Blätter): {dlg.FileName}; Zeilen={_importedRules.Count}");
            TrySuggestPrismaV();
        }
        catch (Exception ex)
        {
            AppendVisualReport($"Regel-Import fehlgeschlagen: {ex.Message}");
            MessageBox.Show(ex.ToString(), "Fehler beim Regel-Import", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Background);
    }

    private async void Analyze_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            StatusText.Text = "Analysiere...";
            var step = StepPathText.Text.Trim();

            if (string.IsNullOrWhiteSpace(step) || !File.Exists(step))
                throw new InvalidOperationException("Bitte STEP-Datei auswaehlen.");
            if (_importedRules.Count == 0)
                throw new InvalidOperationException("Bitte Regeln importieren (Button 'Datei waehlen...').");

            AppendVisualReport("Analyse-Start");

            if (_lastProbeResult is null || !_lastProbeResult.Success)
            {
                App.MarkOperation("Analyze/ProbeOnly");
                _lastProbeResult = await ProbeStepInIsolatedProcessAsync(step, CancellationToken.None);
                AppendVisualReport($"Diagnose: Probe(Analyse) Parts={_lastProbeResult.Parts}, Vertices={_lastProbeResult.Vertices}, Triangles={_lastProbeResult.Triangles}");
                if (!_lastProbeResult.Success)
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(_lastProbeResult.Error) ? "STEP-Probe fehlgeschlagen." : _lastProbeResult.Error);
            }

            if (_lastProbeResult.Thickness is decimal thicknessFromProbe && thicknessFromProbe > 0)
            {
                var selectedMaterial = GetSelectedMaterial();
                var normalized = NormalizeThicknessForRules(selectedMaterial, thicknessFromProbe);

                var de = CultureInfo.GetCultureInfo("de-DE");
                ThicknessText.Text = normalized.ToString("0.00", de);
                if (normalized != thicknessFromProbe)
                    AppendVisualReport($"Diagnose: Dicke aus Probe {thicknessFromProbe:0.##} mm auf Regelstärke {normalized:0.##} mm korrigiert ({selectedMaterial}).");
                else
                    AppendVisualReport($"Diagnose: Dicke aus Probe = {ThicknessText.Text} mm");

                TrySuggestPrismaV();
            }
            else
            {
                AppendVisualReport("Diagnose: Keine Dicke aus Probe verfügbar.");
            }

            var material = GetSelectedMaterial();
            var tRaw = ThicknessText.Text.Trim().Replace(",", ".");
            if (!decimal.TryParse(tRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var thicknessValue))
                throw new InvalidOperationException("Dicke (mm) ist ungueltig.");

            var prismaV = PrismaVText.Text.Trim();
            if (prismaV == "") prismaV = null;

            App.MarkOperation("Analyze/RunRules");
            var result = await _svc.AnalyzeAsync(step, _importedRules, material, thicknessValue, prismaV, CancellationToken.None);
            FindingsGrid.ItemsSource = result.Findings;
            StatusText.Text = $"Fertig. Regel: {(result.SelectedRule is null ? "-" : $"{result.SelectedRule.Material} {result.SelectedRule.ThicknessMm} {result.SelectedRule.PrismaV}")}";
            AppendVisualReport($"Analyse fertig. Findings: {result.Findings.Count}");
            RenderFlatPatternPreview(result);
            App.MarkIdle();
        }
        catch (Exception ex)
        {
            StatusText.Text = "Fehler";
            AppendVisualReport($"Analyse-Fehler: {ex}");
            App.MarkOperation("Fault: Analyse");
            MessageBox.Show(ex.ToString(), "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ExportGeoFromStep_Click(object sender, RoutedEventArgs e)
    {
        var step = StepPathText.Text.Trim();
        if (string.IsNullOrWhiteSpace(step) || !File.Exists(step))
        {
            MessageBox.Show("Bitte zuerst eine STEP-Datei wählen.", "Export GEO", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        StatusText.Text = "Exportiere GEO...";
        AppendVisualReport($"Export GEO Start: {Path.GetFileName(step)}");
        App.MarkOperation("ExportGEO");

        try
        {
            var geoPath = await Task.Run(() =>
            {
                var exporter = new StepToGeoExporter();
                return exporter.Export(step, CancellationToken.None);
            });

            StatusText.Text = $"GEO exportiert: {Path.GetFileName(geoPath)}";
            AppendVisualReport($"GEO gespeichert: {geoPath}");
            App.MarkIdle();

            MessageBox.Show(
                $"GEO erfolgreich erstellt:{Environment.NewLine}{geoPath}",
                "Export GEO",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            StatusText.Text = "GEO-Export fehlgeschlagen.";
            AppendVisualReport($"GEO-Exportfehler: {ex}");
            App.MarkOperation("Fault: ExportGEO");
            MessageBox.Show(ex.ToString(), "GEO-Export Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ResetViewport()
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

        stepViewport.Visibility = Visibility.Visible;
        stepViewport.IsHitTestVisible = true;
        PreviewSnapshotImage.Source = null;
        PreviewSnapshotImage.Visibility = Visibility.Collapsed;

        stepViewport.Children.Clear();
        stepViewport.Children.Add(new SunLight());
        stepViewport.Children.Add(new CoordinateSystemVisual3D());
    }

    private async Task LoadStepIntoUiAsync(string stepPath, CancellationToken ct)
    {
        StatusText.Text = "Lese STEP…";
        AppendVisualReport($"Vorschau-Start: {Path.GetFileName(stepPath)}");

        if (!File.Exists(stepPath))
            throw new FileNotFoundException("STEP-Datei nicht gefunden.", stepPath);

        App.MarkOperation("Probe/ExternalProcess");
        var probe = await ProbeStepInIsolatedProcessAsync(stepPath, ct);
        _lastProbeResult = probe;

        AppendVisualReport($"Diagnose: Parts={probe.Parts}, Vertices={probe.Vertices}, Triangles={probe.Triangles}");
        if (probe.Thickness is decimal t && t > 0)
        {
            var material = GetSelectedMaterial();
            var normalized = NormalizeThicknessForRules(material, t);

            var de = CultureInfo.GetCultureInfo("de-DE");
            ThicknessText.Text = normalized.ToString("0.00", de);
            if (normalized != t)
                AppendVisualReport($"Diagnose: Probe-Dicke {t:0.##} mm auf Regelstärke {normalized:0.##} mm korrigiert ({material}).");
            else
                AppendVisualReport($"Diagnose: Probe-Dicke={ThicknessText.Text} mm");

            TrySuggestPrismaV();
        }
        if (!string.IsNullOrWhiteSpace(probe.Error))
            AppendVisualReport($"Diagnose: Probe-Fehler: {probe.Error}");

        if (!probe.Success)
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(probe.Error) ? "STEP-Probe fehlgeschlagen." : probe.Error);

        if (DisableStepMeshRendering)
        {
            ShowStartupPreview();
            StatusText.Text = "STEP geprüft. Stabiler Prüfmodus aktiv.";
            AppendVisualReport("Diagnose: Externe STEP-Prüfung erfolgreich. Lokales OCCT deaktiviert.");
            return;
        }

        if (string.IsNullOrWhiteSpace(probe.SceneFile) || !File.Exists(probe.SceneFile))
            throw new InvalidOperationException("STEP-Szenedatei aus Probe fehlt.");

        var sceneJson = await File.ReadAllTextAsync(probe.SceneFile, ct);
        var scene = JsonSerializer.Deserialize<StepScene>(sceneJson);
        if (scene is null || scene.Parts.Count == 0)
            throw new InvalidOperationException("Probe lieferte keine darstellbare Szene.");

        _lastSceneForFlatPattern = scene;

        if (!RenderScene(scene))
            throw new InvalidOperationException("In der STEP-Datei wurde keine renderbare Vorschau-Geometrie gefunden.");

        if (_importedRules.Count > 0)
        {
            var material = GetSelectedMaterial();
            var thicknessRaw = ThicknessText.Text.Trim().Replace(",", ".");
            _ = decimal.TryParse(thicknessRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var thicknessMm);
            var prismaV = string.IsNullOrWhiteSpace(PrismaVText.Text) ? null : PrismaVText.Text.Trim();
            var selected = _ruleService.FindBestRule(_importedRules, material, thicknessMm, prismaV);
            var previewResult = new AnalysisResult(stepPath, "ImportedRulesDb", material, thicknessMm, selected, Array.Empty<Finding>());
            RenderFlatPatternPreview(previewResult);
        }
        else
        {
            FlatPatternInfoText.Text = "Regeln importieren, um Abwicklung mit Maßen zu sehen.";
        }

        StatusText.Text = "STEP geladen. Vorschau aktiv.";
        AppendVisualReport("Diagnose: Szene aus externem Prozess gerendert.");
    }

    private async Task<StepProbeResult> ProbeStepInIsolatedProcessAsync(string stepPath, CancellationToken ct)
    {
        var exe = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(exe) || !File.Exists(exe))
            throw new InvalidOperationException("Pfad zur App-EXE konnte nicht ermittelt werden.");

        var outputPath = Path.Combine(Path.GetTempPath(), $"bendchecker-step-probe-{Guid.NewGuid():N}.txt");

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = exe,
                Arguments = $"--step-probe \"{stepPath}\" \"{outputPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true
            };
            psi.Environment["BENDCHECKER_DISABLE_NATIVE_PROBE"] = "1";

            using var process = Process.Start(psi) ?? throw new InvalidOperationException("STEP-Prozess konnte nicht gestartet werden.");
            await process.WaitForExitAsync(ct);

            if (!File.Exists(outputPath))
            {
                var diagPath = Path.Combine(AppContext.BaseDirectory, "diagnostics.txt");
                var details = File.Exists(diagPath) ? $" Diagnostics: {diagPath}" : string.Empty;
                throw new InvalidOperationException($"STEP-Probe lieferte keine Ergebnisdatei. ExitCode={process.ExitCode}.{details}");
            }

            return ParseProbeResult(File.ReadAllLines(outputPath));
        }
        finally
        {
            try
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
            catch
            {
                // ignore cleanup errors
            }
        }
    }

    private static StepProbeResult ParseProbeResult(string[] lines)
    {
        string Get(string key)
        {
            var line = lines.FirstOrDefault(l => l.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase));
            return line is null ? string.Empty : line[(key.Length + 1)..];
        }

        var successRaw = Get("Success");
        var partsRaw = Get("Parts");
        var verticesRaw = Get("Vertices");
        var trianglesRaw = Get("Triangles");
        var thicknessRaw = Get("Thickness");
        var sceneFile = Get("SceneFile");
        var error = Get("Error");

        _ = int.TryParse(partsRaw, out var parts);
        _ = int.TryParse(verticesRaw, out var vertices);
        _ = int.TryParse(trianglesRaw, out var triangles);

        decimal? thickness = null;
        if (decimal.TryParse(thicknessRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var t) && t > 0)
            thickness = t;

        var success = successRaw == "1" || successRaw.Equals("true", StringComparison.OrdinalIgnoreCase);
        return new StepProbeResult(success, parts, vertices, triangles, thickness, sceneFile, error);
    }

    private sealed record StepProbeResult(bool Success, int Parts, int Vertices, int Triangles, decimal? Thickness, string SceneFile, string Error);

    private void TrySuggestPrismaV()
    {
        if (_importedRules.Count == 0)
            return;

        var thicknessRaw = ThicknessText.Text.Trim().Replace(",", ".");
        if (!decimal.TryParse(thicknessRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var thickness))
            return;

        var material = GetSelectedMaterial();
        var v = _ruleService.SuggestPrismaV(_importedRules, material, thickness);
        if (!string.IsNullOrWhiteSpace(v))
        {
            PrismaVText.Text = v;
            StatusText.Text += $", Prisma V: {v}";
            AppendVisualReport($"Prisma V vorgeschlagen: {v}");
        }
    }

    private decimal NormalizeThicknessForRules(string material, decimal measuredThickness)
    {
        if (_importedRules.Count == 0)
            return measuredThickness;

        var candidates = _importedRules
            .Where(r => string.Equals(r.Material, material, StringComparison.OrdinalIgnoreCase))
            .Select(r => r.ThicknessMm)
            .Distinct()
            .OrderBy(v => v)
            .ToList();

        if (candidates.Count == 0)
            return measuredThickness;

        var nearest = candidates.OrderBy(v => Math.Abs(v - measuredThickness)).First();
        var diff = Math.Abs(nearest - measuredThickness);

        // OCCT-Dickenmessung kann bei Blechteilen halbe Werte liefern; nahe Regelstärke bevorzugen.
        return diff <= 0.60m ? nearest : measuredThickness;
    }

    private void ShowStartupPreview()
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

        stepViewport.Visibility = Visibility.Visible;
        stepViewport.IsHitTestVisible = true;
        PreviewSnapshotImage.Source = null;
        PreviewSnapshotImage.Visibility = Visibility.Collapsed;

        stepViewport.Children.Clear();
        stepViewport.Children.Add(new SunLight());
        stepViewport.Children.Add(new CoordinateSystemVisual3D());
        stepViewport.Children.Add(new BoxVisual3D
        {
            Center = new Point3D(0, 0, 0),
            Width = 40,
            Length = 20,
            Height = 10,
            Fill = new SolidColorBrush(Color.FromRgb(185, 190, 205))
        });

        SetCamera(new Rect3D(-20, -10, -5, 40, 20, 10));
        StatusText.Text = "Testmodell sichtbar – 3D-Vorschau aktiv.";
    }

    private bool RenderScene(StepScene scene)
    {
        try
        {
            var stepViewport = EnsureViewport();
            if (stepViewport is null)
                return false;

            stepViewport.Children.Clear();
            stepViewport.Children.Add(new SunLight());
            stepViewport.Children.Add(new CoordinateSystemVisual3D());

            var sourceBounds = Rect3D.Empty;
            foreach (var part in scene.Parts)
            {
                if (part.Positions.Length < 3)
                    continue;

                for (var i = 0; i <= part.Positions.Length - 3; i += 3)
                {
                    var x = part.Positions[i];
                    var y = part.Positions[i + 1];
                    var z = part.Positions[i + 2];
                    if (!IsFinite(x) || !IsFinite(y) || !IsFinite(z))
                        continue;

                    var p = new Point3D(x, y, z);
                    sourceBounds = sourceBounds.IsEmpty ? new Rect3D(p, new Size3D()) : Rect3D.Union(sourceBounds, new Rect3D(p, new Size3D()));
                }
            }

            if (sourceBounds.IsEmpty)
            {
                ShowStartupPreview();
                AppendVisualReport("Render-Info: SourceBounds leer.");
                return false;
            }

            var sourceCenter = new Point3D(
                sourceBounds.X + sourceBounds.SizeX / 2d,
                sourceBounds.Y + sourceBounds.SizeY / 2d,
                sourceBounds.Z + sourceBounds.SizeZ / 2d);

            var bounds = Rect3D.Empty;
            var triangleCount = 0;
            var skippedTriangles = 0;

            foreach (var part in scene.Parts)
            {
                if (part.Positions.Length < 3 || part.Indices.Length < 3)
                    continue;

                var mesh = new MeshGeometry3D();

                for (var i = 0; i <= part.Indices.Length - 3; i += 3)
                {
                    var i1 = part.Indices[i];
                    var i2 = part.Indices[i + 1];
                    var i3 = part.Indices[i + 2];

                    if (!TryReadPosition(part.Positions, i1, out var p1) ||
                        !TryReadPosition(part.Positions, i2, out var p2) ||
                        !TryReadPosition(part.Positions, i3, out var p3))
                    {
                        skippedTriangles++;
                        continue;
                    }

                    if (!IsFinitePoint(p1) || !IsFinitePoint(p2) || !IsFinitePoint(p3))
                    {
                        skippedTriangles++;
                        continue;
                    }

                    var n = TryReadNormal(part.Normals, i1, out var n1) &&
                            TryReadNormal(part.Normals, i2, out var n2) &&
                            TryReadNormal(part.Normals, i3, out var n3) &&
                            IsFiniteVector(n1) && IsFiniteVector(n2) && IsFiniteVector(n3)
                        ? new Vector3D((n1.X + n2.X + n3.X) / 3d, (n1.Y + n2.Y + n3.Y) / 3d, (n1.Z + n2.Z + n3.Z) / 3d)
                        : CalculateNormal(p1, p2, p3);

                    var baseIndex = mesh.Positions.Count;
                    mesh.Positions.Add(new Point3D(p1.X - sourceCenter.X, p1.Y - sourceCenter.Y, p1.Z - sourceCenter.Z));
                    mesh.Positions.Add(new Point3D(p2.X - sourceCenter.X, p2.Y - sourceCenter.Y, p2.Z - sourceCenter.Z));
                    mesh.Positions.Add(new Point3D(p3.X - sourceCenter.X, p3.Y - sourceCenter.Y, p3.Z - sourceCenter.Z));

                    mesh.Normals.Add(n);
                    mesh.Normals.Add(n);
                    mesh.Normals.Add(n);

                    mesh.TriangleIndices.Add(baseIndex);
                    mesh.TriangleIndices.Add(baseIndex + 1);
                    mesh.TriangleIndices.Add(baseIndex + 2);
                }

                if (mesh.TriangleIndices.Count == 0)
                    continue;

                triangleCount += mesh.TriangleIndices.Count / 3;
                bounds = bounds.IsEmpty ? mesh.Bounds : Rect3D.Union(bounds, mesh.Bounds);

                var baseColor = Color.FromArgb(part.Alpha, part.Red, part.Green, part.Blue);
                var material = CreateMaterial(baseColor);
                var model = new GeometryModel3D
                {
                    Geometry = mesh,
                    Material = material,
                    BackMaterial = material
                };

                stepViewport.Children.Add(new ModelVisual3D { Content = model });
            }

            if (triangleCount <= 0 || bounds.IsEmpty)
            {
                ShowStartupPreview();
                AppendVisualReport("Render-Info: Keine gültigen Dreiecke oder Bounds leer.");
                return false;
            }

            SetCamera(bounds);
            stepViewport.ZoomExtents(0);

            if (PreviewSafeMode)
            {
                stepViewport.UpdateLayout();
                var w = Math.Max(1, (int)Math.Round(stepViewport.ActualWidth));
                var h = Math.Max(1, (int)Math.Round(stepViewport.ActualHeight));
                if (w > 1 && h > 1)
                {
                    var bmp = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
                    bmp.Render(stepViewport);
                    bmp.Freeze();

                    PreviewSnapshotImage.Source = bmp;
                    PreviewSnapshotImage.Visibility = Visibility.Visible;
                }

                // Live-Viewport vollständig deaktivieren/freigeben, um spätere native Render-Crashes zu vermeiden.
                stepViewport.Children.Clear();
                PreviewHost.Children.Remove(stepViewport);
                _stepViewport = null;
                AppendVisualReport("Diagnose: Safe-Mode aktiv (statisches Vorschau-Bild, Live-3D entladen).");
            }

            StatusText.Text = $"STEP-Vorschau geladen: {triangleCount} Dreiecke.";
            AppendVisualReport($"Render ok. Dreiecke: {triangleCount}; Skipped: {skippedTriangles}; Bounds: {bounds}");
            return true;
        }
        catch (Exception ex)
        {
            ShowStartupPreview();
            AppendVisualReport($"Render-Fehler: {ex}");
            MessageBox.Show(ex.ToString(), "Render-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    private static bool TryReadPosition(double[] positions, int index, out Point3D point)
    {
        var offset = index * 3;
        if (index < 0 || offset < 0 || offset + 2 >= positions.Length)
        {
            point = default;
            return false;
        }

        point = new Point3D(positions[offset], positions[offset + 1], positions[offset + 2]);
        return true;
    }

    private static bool TryReadNormal(double[] normals, int index, out Vector3D normal)
    {
        var offset = index * 3;
        if (index < 0 || offset < 0 || offset + 2 >= normals.Length)
        {
            normal = default;
            return false;
        }

        normal = new Vector3D(normals[offset], normals[offset + 1], normals[offset + 2]);
        return true;
    }

    private static Vector3D CalculateNormal(Point3D p1, Point3D p2, Point3D p3)
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
            return new Vector3D(0, 0, 1);

        return new Vector3D(nx / length, ny / length, nz / length);
    }

    private static bool IsFinitePoint(Point3D p) => IsFinite(p.X) && IsFinite(p.Y) && IsFinite(p.Z);

    private static bool IsFiniteVector(Vector3D v) => IsFinite(v.X) && IsFinite(v.Y) && IsFinite(v.Z);

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static Material CreateMaterial(Color color)
    {
        var diffuseBrush = new SolidColorBrush(color);
        var emissiveBrush = new SolidColorBrush(Color.FromArgb(
            color.A,
            (byte)Math.Max(18, color.R / 4),
            (byte)Math.Max(18, color.G / 4),
            (byte)Math.Max(18, color.B / 4)));

        var material = new MaterialGroup();
        material.Children.Add(new EmissiveMaterial(emissiveBrush));
        material.Children.Add(new DiffuseMaterial(diffuseBrush));
        return material;
    }

    private void SetCamera(Rect3D bounds)
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

        var center = new Point3D(
            bounds.X + bounds.SizeX / 2d,
            bounds.Y + bounds.SizeY / 2d,
            bounds.Z + bounds.SizeZ / 2d);

        var maxSize = Math.Max(Math.Max(bounds.SizeX, bounds.SizeY), bounds.SizeZ);
        var radius = Math.Max(maxSize / 2d, 0.5d);
        var distance = Math.Max(radius * 4d, 8d);

        stepViewport.Camera = new PerspectiveCamera
        {
            Position = new Point3D(center.X + distance, center.Y + distance, center.Z + distance),
            LookDirection = new Vector3D(-distance, -distance, -distance),
            UpDirection = new Vector3D(0, 0, 1),
            FieldOfView = 45,
            NearPlaneDistance = Math.Max(distance / 500d, 0.01d),
            FarPlaneDistance = Math.Max(distance * 500d, 1000d)
        };
    }

    private HelixViewport3D? EnsureViewport()
    {
        if (_isDesignMode)
            return null;

        if (_stepViewport is not null)
            return _stepViewport;

        DesignerPreviewPlaceholder.Visibility = Visibility.Collapsed;
        _stepViewport = new HelixViewport3D
        {
            Background = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
            ShowCoordinateSystem = true,
            ShowViewCube = true,
            ZoomExtentsWhenLoaded = false,
            RotateAroundMouseDownPoint = true,
            IsHeadLightEnabled = true
        };

        PreviewHost.Children.Add(_stepViewport);
        return _stepViewport;
    }

    public void ReportFromApp(string source, string details)
    {
        AppendVisualReport($"{source}: {details}");

        if (source.Contains("Exception", StringComparison.OrdinalIgnoreCase) || source.Contains("Fault", StringComparison.OrdinalIgnoreCase))
        {
            StatusText.Text = "Kritischer Fehler erkannt. Siehe Report.";
            App.MarkOperation($"Fault event: {source}");
        }
    }

    private void AppendVisualReport(string message)
    {
        if (!Dispatcher.CheckAccess())
        {
            _ = Dispatcher.BeginInvoke(() => AppendVisualReport(message));
            return;
        }

        var compact = (message ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
        if (compact.Length > MaxReportEntryLength)
            compact = compact[..MaxReportEntryLength] + " ...";

        var line = $"[{DateTime.Now:HH:mm:ss}] {compact}";
        _reportLines.Add(line);
        if (_reportLines.Count > MaxReportLines)
            _reportLines.RemoveRange(0, _reportLines.Count - MaxReportLines);

        VisualizationReportText.Text = string.Join(Environment.NewLine, _reportLines);
        VisualizationReportText.ScrollToEnd();

        App.AppendActivityLog(compact);
    }

    private void SelectAllReport_Click(object sender, RoutedEventArgs e)
    {
        VisualizationReportText.Focus();
        VisualizationReportText.SelectAll();
    }

    private void CopyReport_Click(object sender, RoutedEventArgs e)
    {
        var text = string.Join(Environment.NewLine, _reportLines);
        if (string.IsNullOrWhiteSpace(text))
        {
            StatusText.Text = "Report ist leer.";
            return;
        }

        try
        {
            Clipboard.Clear();
            Clipboard.SetDataObject(text, true);
            StatusText.Text = "Report in Zwischenablage kopiert.";
        }
        catch (Exception ex)
        {
            AppendVisualReport($"Kopierfehler: {ex.Message}");
            StatusText.Text = "Kopieren fehlgeschlagen. Bitte 'Report speichern' nutzen.";
            MessageBox.Show(ex.ToString(), "Kopierfehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveReport_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var text = string.Join(Environment.NewLine, _reportLines);
            if (string.IsNullOrWhiteSpace(text))
            {
                StatusText.Text = "Report ist leer.";
                return;
            }

            var path = App.SaveTextReport(text);
            StatusText.Text = "Report gespeichert.";
            AppendVisualReport($"Report gespeichert: {path}");
        }
        catch (Exception ex)
        {
            AppendVisualReport($"Speicherfehler: {ex.Message}");
            StatusText.Text = "Report konnte nicht gespeichert werden.";
            MessageBox.Show(ex.ToString(), "Speicherfehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenReportFolder_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var folder = App.GetReportFolderPath();
            Process.Start(new ProcessStartInfo
            {
                FileName = folder,
                UseShellExecute = true
            });
            StatusText.Text = "Report-Ordner geöffnet.";
        }
        catch (Exception ex)
        {
            AppendVisualReport($"Ordner öffnen fehlgeschlagen: {ex.Message}");
            StatusText.Text = "Report-Ordner konnte nicht geöffnet werden.";
            MessageBox.Show(ex.ToString(), "Ordner öffnen fehlgeschlagen", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void TryRestoreRulesPath()
    {
        try
        {
            if (!File.Exists(_savedRulesPathFile))
            {
                if (_importedRules.Count == 0)
                    RulesInfoText.Text = "Keine Regeln importiert.";
                return;
            }

            var saved = File.ReadAllText(_savedRulesPathFile).Trim();
            if (string.IsNullOrWhiteSpace(saved) || !File.Exists(saved))
            {
                RulesInfoText.Text = _importedRules.Count == 0 ? "Keine Regeln importiert." : $"Regeln geladen: {_importedRules.Count} Zeilen";
                return;
            }

            RulesPathText.Text = saved;
            RulesInfoText.Text = _importedRules.Count == 0 ? "Regeln nicht importiert." : $"Regeln geladen: {_importedRules.Count} Zeilen";
            AppendVisualReport($"Regeldatei automatisch gesetzt: {saved}");
            TrySuggestPrismaV();
        }
        catch (Exception ex)
        {
            AppendVisualReport($"Regelpfad konnte nicht wiederhergestellt werden: {ex.Message}");
            MessageBox.Show(ex.ToString(), "Regelpfad-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveRulesPath(string path)
    {
        try
        {
            File.WriteAllText(_savedRulesPathFile, path ?? string.Empty);
        }
        catch (Exception ex)
        {
            AppendVisualReport($"Regelpfad konnte nicht gespeichert werden: {ex.Message}");
            MessageBox.Show(ex.ToString(), "Regelpfad-Speicherfehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadRulesDb()
    {
        try
        {
            if (!File.Exists(_rulesDbFile))
            {
                _importedRules = [];
                RulesInfoText.Text = "Keine Regeln importiert.";
                return;
            }

            var json = File.ReadAllText(_rulesDbFile);
            _importedRules = JsonSerializer.Deserialize<List<RuleRow>>(json) ?? [];
            RulesInfoText.Text = _importedRules.Count == 0 ? "Keine Regeln importiert." : $"Regeln geladen: {_importedRules.Count} Zeilen";
            AppendVisualReport($"Regel-Datenbank geladen: {_importedRules.Count} Zeilen");
        }
        catch (Exception ex)
        {
            _importedRules = [];
            RulesInfoText.Text = "Regel-DB fehlerhaft.";
            AppendVisualReport($"Regel-DB konnte nicht geladen werden: {ex.Message}");
            MessageBox.Show(ex.ToString(), "Regel-DB Ladefehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveRulesDb()
    {
        try
        {
            var json = JsonSerializer.Serialize(_importedRules);
            File.WriteAllText(_rulesDbFile, json);
        }
        catch (Exception ex)
        {
            AppendVisualReport($"Regel-DB konnte nicht gespeichert werden: {ex.Message}");
            MessageBox.Show(ex.ToString(), "Regel-DB Speicherfehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private string GetSelectedMaterial()
    {
        return ((System.Windows.Controls.ComboBoxItem)MaterialBox.SelectedItem).Content?.ToString() ?? "Stahl";
    }

    private void RenderFlatPatternPreview(AnalysisResult result)
    {
        // 1) Falls echte GEO importiert wurde, diese bevorzugen.
        if (_hasImportedRealGeo)
        {
            FlatPatternImage.Source = null;
            FlatPatternInfoText.Text = "Echte GEO geladen. Vorschau über '.geo Vorschau'.";
            return;
        }

        // 2) STEP->GEO (Beta): projizierte Konturschleifen aus Szene.
        if (TryBuildProjectedLoopsFromLastScene(out var loopsMm, out var outlineWidthMm, out var outlineHeightMm))
        {
            _lastFlatPatternGeo = BuildGeoFromLoops(loopsMm, result, outlineWidthMm, outlineHeightMm);
            _hasGeneratedGeo = !string.IsNullOrWhiteSpace(_lastFlatPatternGeo);

            FlatPatternInfoText.Text = $"Beta-Abwicklung aus STEP: {outlineWidthMm:0.##} x {outlineHeightMm:0.##} mm, Loops={loopsMm.Count}";
            AppendVisualReport($"Beta-GEO aus STEP erzeugt. Loops={loopsMm.Count}, Größe={outlineWidthMm:0.##}x{outlineHeightMm:0.##} mm");
            return;
        }

        // 3) Fallback: Außenkontur als Hull
        if (TryBuildProjectedOutlineFromLastScene(out var outlineMm, out var wMm, out var hMm))
        {
            _lastFlatPatternGeo = BuildGeoFromLoops([outlineMm], result, wMm, hMm);
            _hasGeneratedGeo = !string.IsNullOrWhiteSpace(_lastFlatPatternGeo);
            FlatPatternInfoText.Text = $"Beta-Abwicklung (Fallback): {wMm:0.##} x {hMm:0.##} mm";
            AppendVisualReport($"Beta-GEO Fallback erzeugt. Größe={wMm:0.##}x{hMm:0.##} mm");
            return;
        }

        FlatPatternImage.Source = null;
        FlatPatternInfoText.Text = "Keine STEP-Abwicklung erzeugbar. Optional echte GEO laden.";
        _lastFlatPatternGeo = string.Empty;
        _hasGeneratedGeo = false;
        AppendVisualReport("Abwicklung: Aus STEP konnte keine GEO erzeugt werden.");
    }

    private bool TryBuildProjectedLoopsFromLastScene(out List<List<Point>> loopsMm, out double widthMm, out double heightMm)
    {
        loopsMm = [];
        widthMm = 0;
        heightMm = 0;

        if (!TryBuildProjectedOutlineFromLastScene(out var outline, out widthMm, out heightMm))
            return false;

        loopsMm = [outline];
        return true;
    }

    private bool TryBuildProjectedOutlineFromLastScene(out List<Point> outlineMm, out double widthMm, out double heightMm)
    {
        outlineMm = [];
        widthMm = 0d;
        heightMm = 0d;

        if (_lastSceneForFlatPattern is null || _lastSceneForFlatPattern.Parts.Count == 0)
            return false;

        var xs = new List<double>();
        var ys = new List<double>();
        var zs = new List<double>();

        foreach (var part in _lastSceneForFlatPattern.Parts)
        {
            for (var i = 0; i <= part.Positions.Length - 3; i += 3)
            {
                var x = part.Positions[i];
                var y = part.Positions[i + 1];
                var z = part.Positions[i + 2];
                if (!IsFinite(x) || !IsFinite(y) || !IsFinite(z))
                    continue;

                xs.Add(x);
                ys.Add(y);
                zs.Add(z);
            }
        }

        if (xs.Count < 20)
            return false;

        var rx = xs.Max() - xs.Min();
        var ry = ys.Max() - ys.Min();
        var rz = zs.Max() - zs.Min();

        var plane = SelectBestProjectionPlane(rx, ry, rz);
        List<Point> projected = plane switch
        {
            ProjectionPlane.XY => xs.Zip(ys, (a, b) => new Point(a, b)).ToList(),
            ProjectionPlane.XZ => xs.Zip(zs, (a, b) => new Point(a, b)).ToList(),
            _ => ys.Zip(zs, (a, b) => new Point(a, b)).ToList()
        };

        var hull = ComputeConvexHull(projected);
        if (hull.Count < 3)
            return false;

        var minX = hull.Min(p => p.X);
        var minY = hull.Min(p => p.Y);
        outlineMm = hull.Select(p => new Point(p.X - minX, p.Y - minY)).ToList();

        widthMm = outlineMm.Max(p => p.X);
        heightMm = outlineMm.Max(p => p.Y);
        return widthMm > 0.1 && heightMm > 0.1;
    }

    private string BuildGeoFromLoops(
        IReadOnlyList<IReadOnlyList<Point>> loopsMm,
        AnalysisResult result,
        double widthMm,
        double heightMm)
    {
        if (loopsMm.Count == 0 || loopsMm[0].Count < 3)
            return string.Empty;

        var material = GetSelectedMaterial();
        var thicknessText = ThicknessText.Text.Trim().Replace(",", ".");
        _ = double.TryParse(thicknessText, NumberStyles.Any, CultureInfo.InvariantCulture, out var thicknessMm);
        var prismaV = string.IsNullOrWhiteSpace(PrismaVText.Text) ? "-" : PrismaVText.Text.Trim();

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("# BendChecker Flat Pattern GEO");
        sb.AppendLine("$Units mm");
        sb.AppendLine("$Type Polygon2D");
        sb.AppendLine("$Process PressBrake");
        sb.AppendLine($"$Material {material}");
        sb.AppendLine($"$Thickness {thicknessMm.ToString("0.###", CultureInfo.InvariantCulture)}");
        sb.AppendLine($"$PrismaV {prismaV}");
        sb.AppendLine($"$BoundingWidth {widthMm.ToString("0.###", CultureInfo.InvariantCulture)}");
        sb.AppendLine($"$BoundingHeight {heightMm.ToString("0.###", CultureInfo.InvariantCulture)}");
        sb.AppendLine($"$LoopCount {loopsMm.Count}");

        var pointNames = new List<List<string>>();
        var pointIndex = 1;
        for (var li = 0; li < loopsMm.Count; li++)
        {
            var names = new List<string>();
            foreach (var p in loopsMm[li])
            {
                var name = $"P{pointIndex++}";
                names.Add(name);
                sb.AppendLine($"{name} {p.X.ToString("0.###", CultureInfo.InvariantCulture)} {p.Y.ToString("0.###", CultureInfo.InvariantCulture)}");
            }
            pointNames.Add(names);
        }

        for (var li = 0; li < pointNames.Count; li++)
            sb.AppendLine($"LOOP{li + 1} {string.Join(" ", pointNames[li])}");

        if (result.SelectedRule is not null)
            sb.AppendLine($"$Rule {result.SelectedRule.Material}|t={result.SelectedRule.ThicknessMm:0.###}|V={result.SelectedRule.PrismaV}");

        return sb.ToString();
    }

    private enum ProjectionPlane { XY, XZ, YZ }

    private static ProjectionPlane SelectBestProjectionPlane(double rx, double ry, double rz)
    {
        var areaXY = rx * ry;
        var areaXZ = rx * rz;
        var areaYZ = ry * rz;
        if (areaXY >= areaXZ && areaXY >= areaYZ) return ProjectionPlane.XY;
        if (areaXZ >= areaXY && areaXZ >= areaYZ) return ProjectionPlane.XZ;
        return ProjectionPlane.YZ;
    }

    private static List<Point> ComputeConvexHull(List<Point> points)
    {
        var pts = points.OrderBy(p => p.X).ThenBy(p => p.Y).Distinct().ToList();
        if (pts.Count <= 3) return pts;

        static double Cross(Point o, Point a, Point b) => (a.X - o.X) * (b.Y - o.Y) - (a.Y - o.Y) * (b.X - o.X);

        var lower = new List<Point>();
        foreach (var p in pts)
        {
            while (lower.Count >= 2 && Cross(lower[^2], lower[^1], p) <= 0) lower.RemoveAt(lower.Count - 1);
            lower.Add(p);
        }

        var upper = new List<Point>();
        for (var i = pts.Count - 1; i >= 0; i--)
        {
            var p = pts[i];
            while (upper.Count >= 2 && Cross(upper[^2], upper[^1], p) <= 0) upper.RemoveAt(upper.Count - 1);
            upper.Add(p);
        }

        lower.RemoveAt(lower.Count - 1);
        upper.RemoveAt(upper.Count - 1);
        lower.AddRange(upper);
        return lower;
    }

    private async Task<string> LoadGeoFileAsync(string path, CancellationToken ct)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException("GEO-Datei nicht gefunden.", path);

        AppendVisualReport($"Lade GEO-Definition: {path}");
        var text = await File.ReadAllTextAsync(path, ct);

        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException("GEO-Datei ist leer.");

        return text;
    }

    private void ExportFlatPatternGeo_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_lastFlatPatternGeo) || (!_hasImportedRealGeo && !_hasGeneratedGeo))
            {
                StatusText.Text = "Keine GEO zum Export vorhanden.";
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "Geo-Datei (*.geo)|*.geo|Alle Dateien (*.*)|*.*",
                FileName = $"Abwicklung-{DateTime.Now:yyyyMMdd-HHmmss}.geo"
            };

            if (dlg.ShowDialog() != true)
                return;

            File.WriteAllText(dlg.FileName, _lastFlatPatternGeo);
            StatusText.Text = "GEO gespeichert.";
            AppendVisualReport($"GEO exportiert: {dlg.FileName}");
        }
        catch (Exception ex)
        {
            StatusText.Text = "Geo-Export fehlgeschlagen.";
            AppendVisualReport($"Geo-Exportfehler: {ex.Message}");
            MessageBox.Show(ex.ToString(), "Geo-Exportfehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadGeoPreview_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Geo-Datei (*.geo)|*.geo|Alle Dateien (*.*)|*.*"
            };

            if (dlg.ShowDialog() != true)
                return;

            var text = File.ReadAllText(dlg.FileName);
            if (string.IsNullOrWhiteSpace(text))
                throw new InvalidOperationException("GEO-Datei ist leer.");

            _lastFlatPatternGeo = text;
            _hasImportedRealGeo = true;
            _hasGeneratedGeo = false;
            FlatPatternInfoText.Text = $"Echte GEO geladen: {Path.GetFileName(dlg.FileName)}";
            AppendVisualReport($"GEO geladen: {dlg.FileName}");
            StatusText.Text = "GEO geladen.";

            var wnd = new GeoPreviewWindow(_lastFlatPatternGeo)
            {
                Owner = this
            };
            wnd.ShowDialog();
        }
        catch (Exception ex)
        {
            StatusText.Text = "GEO laden fehlgeschlagen.";
            AppendVisualReport($"GEO-Ladefehler: {ex.Message}");
            MessageBox.Show(ex.ToString(), "GEO-Ladefehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenGeoPreview_Click(object sender, RoutedEventArgs e)
    {
        if (!_hasImportedRealGeo || string.IsNullOrWhiteSpace(_lastFlatPatternGeo))
        {
            StatusText.Text = "Keine geladene GEO vorhanden.";
            return;
        }

        var wnd = new GeoPreviewWindow(_lastFlatPatternGeo)
        {
            Owner = this
        };
        wnd.ShowDialog();
    }

    private void ZoomIn_Click(object sender, RoutedEventArgs e)
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

        if (stepViewport.Camera is PerspectiveCamera cam)
        {
            cam.Position = new Point3D(cam.Position.X * 0.85, cam.Position.Y * 0.85, cam.Position.Z * 0.85);
            cam.LookDirection = new Vector3D(cam.LookDirection.X * 0.85, cam.LookDirection.Y * 0.85, cam.LookDirection.Z * 0.85);
            StatusText.Text = "Zoom +";
        }
    }

    private void ZoomOut_Click(object sender, RoutedEventArgs e)
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

        if (stepViewport.Camera is PerspectiveCamera cam)
        {
            cam.Position = new Point3D(cam.Position.X * 1.15, cam.Position.Y * 1.15, cam.Position.Z * 1.15);
            cam.LookDirection = new Vector3D(cam.LookDirection.X * 1.15, cam.LookDirection.Y * 1.15, cam.LookDirection.Z * 1.15);
            StatusText.Text = "Zoom -";
        }
    }

    private void ResetView_Click(object sender, RoutedEventArgs e)
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

        stepViewport.ZoomExtents(250);
        StatusText.Text = "Ansicht zurückgesetzt.";
    }
}

