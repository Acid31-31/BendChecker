using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using BendChecker.Core.Models;
using BendChecker.Core.Services;
using HelixToolkit.Wpf;
using Microsoft.Win32;

namespace BendChecker.App;

public partial class MainWindow : Window
{
    private readonly RuleService _ruleService;
    private readonly IStepAnalyzer _stepAnalyzer;
    private readonly IStepPreviewLoader _stepPreviewLoader;
    private readonly BendCheckService _svc;
    private readonly bool _isDesignMode;
    private HelixViewport3D? _stepViewport;

    public MainWindow()
    {
        InitializeComponent();

        _isDesignMode = DesignerProperties.GetIsInDesignMode(this);
        _ruleService = new RuleService();

        if (_isDesignMode)
        {
            _stepAnalyzer = new StepAnalyzerStub();
            _stepPreviewLoader = new DesignTimeStepPreviewLoader();
        }
        else
        {
            var stepAnalyzer = new OcctStepAnalyzer();
            _stepAnalyzer = stepAnalyzer;
            _stepPreviewLoader = stepAnalyzer;
        }

        _svc = new BendCheckService(_ruleService, _stepAnalyzer);

        if (_isDesignMode)
        {
            StatusText.Text = "Designer-Vorschau aktiv.";
            return;
        }

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
        StatusText.Text = "3D-Vorschau bereit.";
    }

    private async void PickStep_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "STEP (*.step;*.stp)|*.step;*.stp|All files (*.*)|*.*" };
        if (dlg.ShowDialog() != true)
            return;

        StepPathText.Text = dlg.FileName;

        try
        {
            StatusText.Text = "Lese STEP…";

            var stepPath = StepPathText.Text.Trim();
            await LoadAndRenderStepAsync(stepPath, CancellationToken.None);

            var thickness = await _stepAnalyzer.TryGetThicknessMmAsync(stepPath, CancellationToken.None);
            if (thickness is not null)
            {
                var de = CultureInfo.GetCultureInfo("de-DE");
                ThicknessText.Text = thickness.Value.ToString("0.00", de);
                StatusText.Text = $"STEP geladen. Dicke erkannt: {ThicknessText.Text} mm";

                var rulesPath = RulesPathText.Text.Trim();
                if (!string.IsNullOrWhiteSpace(rulesPath) && File.Exists(rulesPath))
                {
                    var material = GetSelectedMaterial();
                    var rules = _ruleService.LoadRules(rulesPath);
                    var v = _ruleService.SuggestPrismaV(rules, material, thickness.Value);
                    PrismaVText.Text = v ?? "";
                    if (!string.IsNullOrWhiteSpace(v))
                        StatusText.Text += $", Prisma V: {v}";
                }
            }
            else
            {
                StatusText.Text = "STEP geladen. Dicke nicht gefunden – bitte manuell eingeben.";
            }
        }
        catch (Exception ex)
        {
            ResetViewport();
            ShowStartupPreview();
            StatusText.Text = "Fehler beim STEP lesen";
            MessageBox.Show(ex.Message, "BendChecker", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void PickRules_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*" };
        if (dlg.ShowDialog() == true) RulesPathText.Text = dlg.FileName;
    }

    private async void Analyze_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            StatusText.Text = "Analysiere...";
            var step = StepPathText.Text.Trim();
            var rules = RulesPathText.Text.Trim();

            if (string.IsNullOrWhiteSpace(step) || !File.Exists(step))
                throw new InvalidOperationException("Bitte STEP-Datei auswaehlen.");
            if (string.IsNullOrWhiteSpace(rules) || !File.Exists(rules))
                throw new InvalidOperationException("Bitte Excel-Regeldatei auswaehlen.");

            await LoadAndRenderStepAsync(step, CancellationToken.None);

            var material = GetSelectedMaterial();
            var tRaw = ThicknessText.Text.Trim().Replace(",", ".");
            if (!decimal.TryParse(tRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var thickness))
                throw new InvalidOperationException("Dicke (mm) ist ungueltig.");

            var prismaV = PrismaVText.Text.Trim();
            if (prismaV == "") prismaV = null;

            var result = await _svc.AnalyzeAsync(step, rules, material, thickness, prismaV, CancellationToken.None);
            FindingsGrid.ItemsSource = result.Findings;
            StatusText.Text = $"Fertig. Regel: {(result.SelectedRule is null ? "-" : $"{result.SelectedRule.Material} {result.SelectedRule.ThicknessMm} {result.SelectedRule.PrismaV}")}";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Fehler";
            MessageBox.Show(ex.Message, "BendChecker", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ResetViewport()
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

        stepViewport.Children.Clear();
        stepViewport.Children.Add(new SunLight());
        stepViewport.Children.Add(new CoordinateSystemVisual3D());
    }

    private async Task LoadAndRenderStepAsync(string stepPath, CancellationToken ct)
    {
        var preview = await _stepPreviewLoader.LoadPreviewAsync(stepPath, ct);
        RenderPreview(preview);
    }

    private void ShowStartupPreview()
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

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

    private void RenderPreview(StepPreviewScene scene)
    {
        var mesh = new MeshGeometry3D();
        for (var i = 0; i <= scene.TriangleVertices.Length - 3; i += 3)
        {
            mesh.Positions.Add(new Point3D(scene.TriangleVertices[i], scene.TriangleVertices[i + 1], scene.TriangleVertices[i + 2]));
            mesh.TriangleIndices.Add(mesh.Positions.Count - 1);
        }

        RenderModel(mesh, $"STEP-Vorschau geladen: {mesh.Positions.Count / 3} Dreiecke.");
    }

    private void RenderModel(MeshGeometry3D mesh, string status)
    {
        var stepViewport = EnsureViewport();
        if (stepViewport is null)
            return;

        var brush = new SolidColorBrush(Color.FromRgb(185, 190, 205));
        var material = new MaterialGroup();
        material.Children.Add(new EmissiveMaterial(new SolidColorBrush(Color.FromRgb(42, 46, 56))));
        material.Children.Add(new DiffuseMaterial(brush));

        var model = new GeometryModel3D
        {
            Geometry = mesh,
            Material = material,
            BackMaterial = material
        };

        stepViewport.Children.Clear();
        stepViewport.Children.Add(new SunLight());
        stepViewport.Children.Add(new CoordinateSystemVisual3D());
        stepViewport.Children.Add(new ModelVisual3D { Content = model });

        SetCamera(mesh.Bounds);
        StatusText.Text = status;
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

        var radius = Math.Max(Math.Max(bounds.SizeX, bounds.SizeY), Math.Max(bounds.SizeZ, 1d));
        var distance = Math.Max(radius * 3d, 100d);

        stepViewport.Camera = new PerspectiveCamera
        {
            Position = new Point3D(center.X + distance, center.Y + distance, center.Z + distance),
            LookDirection = new Vector3D(-distance, -distance, -distance),
            UpDirection = new Vector3D(0, 0, 1),
            FieldOfView = 45
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
            ZoomExtentsWhenLoaded = false
        };

        PreviewHost.Children.Add(_stepViewport);
        return _stepViewport;
    }

    private string GetSelectedMaterial()
    {
        return ((System.Windows.Controls.ComboBoxItem)MaterialBox.SelectedItem).Content?.ToString() ?? "Stahl";
    }

    private sealed class DesignTimeStepPreviewLoader : IStepPreviewLoader
    {
        private static readonly StepPreviewScene Scene = new([
            -20d, -10d, 0d,
            20d, -10d, 0d,
            20d, 10d, 0d,
            -20d, -10d, 0d,
            20d, 10d, 0d,
            -20d, 10d, 0d
        ]);

        public Task<StepPreviewScene> LoadPreviewAsync(string stepPath, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            return Task.FromResult(Scene);
        }
    }
}

