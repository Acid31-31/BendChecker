using System.Globalization;
using System.Windows;
using Microsoft.Win32;
using BendChecker.Core.Services;

namespace BendChecker.App;

public partial class MainWindow : Window
{
    private readonly BendCheckService _svc;

    public MainWindow()
    {
        InitializeComponent();

        var ruleService = new RuleService();
        var stepAnalyzer = new StepAnalyzerStub();
        _svc = new BendCheckService(ruleService, stepAnalyzer);
    }

    private void PickStep_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "STEP (*.step;*.stp)|*.step;*.stp|All files (*.*)|*.*"
        };
        if (dlg.ShowDialog() == true)
            StepPathText.Text = dlg.FileName;
    }

    private void PickRules_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        };
        if (dlg.ShowDialog() == true)
            RulesPathText.Text = dlg.FileName;
    }

    private async void Analyze_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            StatusText.Text = "Analysiere…";

            var step = StepPathText.Text.Trim();
            var rules = RulesPathText.Text.Trim();

            if (string.IsNullOrWhiteSpace(step) || !File.Exists(step))
                throw new InvalidOperationException("Bitte STEP-Datei auswählen.");
            if (string.IsNullOrWhiteSpace(rules) || !File.Exists(rules))
                throw new InvalidOperationException("Bitte Excel-Regeldatei auswählen.");

            var material = ((System.Windows.Controls.ComboBoxItem)MaterialBox.SelectedItem).Content?.ToString() ?? "Stahl";

            var tRaw = ThicknessText.Text.Trim().Replace(",", ".");
            if (!decimal.TryParse(tRaw, NumberStyles.Any, CultureInfo.InvariantCulture, out var thickness))
                throw new InvalidOperationException("Dicke (mm) ist ungültig.");

            var prismaV = PrismaVText.Text.Trim();
            if (prismaV == "") prismaV = null;

            var result = await _svc.AnalyzeAsync(step, rules, material, thickness, prismaV, CancellationToken.None);

            FindingsGrid.ItemsSource = result.Findings;
            StatusText.Text = $"Fertig. Regel: {(result.SelectedRule is null ? "—" : $"{result.SelectedRule.Material} {result.SelectedRule.ThicknessMm} {result.SelectedRule.PrismaV}")}";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Fehler";
            MessageBox.Show(ex.Message, "BendChecker", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
