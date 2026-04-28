using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using BendChecker.Core.Models;
using BendChecker.Core.Services;

namespace BendChecker.App;

public partial class App : Application
{
    private static readonly string ReportFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "BendChecker-Reports");
    private static readonly string CrashLogPath = Path.Combine(ReportFolder, "BendChecker-crash.log");
    private static readonly string ActivityLogPath = Path.Combine(ReportFolder, "BendChecker-activity.log");
    private static readonly string SessionStatePath = Path.Combine(ReportFolder, "BendChecker-session.state");
    private static string _startupDiagnostics = "Startup-Diagnose: keine vorherige Sitzung gefunden.";

    protected override void OnStartup(StartupEventArgs e)
    {
        EnsureReportFolder();

        ConfigureOcctRuntime();
        if (TryRunStepProbeMode(e.Args))
        {
            Shutdown();
            return;
        }

        var previous = ReadSessionState();
        if (previous is null)
        {
            _startupDiagnostics = "Startup-Diagnose: keine vorherige Sitzung gefunden.";
        }
        else
        {
            _startupDiagnostics = $"Startup-Diagnose: Vorheriger Status={previous.Value.Status}, Letzte Aktion={previous.Value.Operation}";
            AppendActivityLog(_startupDiagnostics);

            if (string.Equals(previous.Value.Status, "Running", StringComparison.OrdinalIgnoreCase))
            {
                var msg = $"Vorheriger Lauf wurde unerwartet beendet. Letzte Aktion: {previous.Value.Operation}";
                AppendActivityLog(msg);
                MessageBox.Show(msg + $"{Environment.NewLine}Details im Ordner: {ReportFolder}", "Absturz erkannt", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        WriteSessionState("Running", "Startup");
        AppendActivityLog("App-Start");
        TryForceSoftwareRendering();

        RegisterGlobalExceptionHandlers();
        base.OnStartup(e);
    }

    private static bool TryRunStepProbeMode(string[] args)
    {
        if (args.Length < 3 || !string.Equals(args[0], "--step-probe", StringComparison.OrdinalIgnoreCase))
            return false;

        var stepPath = args[1];
        var outputPath = args[2];
        var scenePath = outputPath + ".scene.json";

        try
        {
            var analyzer = new OcctStepAnalyzer();
            var scene = analyzer.TryLoadSceneAsync(stepPath, CancellationToken.None).GetAwaiter().GetResult();
            var thickness = analyzer.TryGetThicknessMmAsync(stepPath, CancellationToken.None).GetAwaiter().GetResult();

            if (scene is not null)
                File.WriteAllText(scenePath, JsonSerializer.Serialize(scene));

            var parts = scene?.Parts.Count ?? 0;
            var vertices = scene?.Parts.Sum(p => p.Positions.Length / 3) ?? 0;
            var triangles = scene?.Parts.Sum(p => p.Indices.Length / 3) ?? 0;

            var lines = new List<string>
            {
                "Success=1",
                $"Parts={parts}",
                $"Vertices={vertices}",
                $"Triangles={triangles}",
                $"Thickness={(thickness is null ? string.Empty : thickness.Value.ToString(System.Globalization.CultureInfo.InvariantCulture))}",
                $"SceneFile={(scene is null ? string.Empty : scenePath)}",
                "Error="
            };
            File.WriteAllLines(outputPath, lines);
            Environment.ExitCode = 0;
        }
        catch (Exception ex)
        {
            var lines = new List<string>
            {
                "Success=0",
                "Parts=0",
                "Vertices=0",
                "Triangles=0",
                "Thickness=",
                "SceneFile=",
                $"Error={ex.GetType().Name}: {ex.Message}"
            };
            try
            {
                File.WriteAllLines(outputPath, lines);
            }
            catch
            {
                // ignore file write failures in probe mode
            }

            Environment.ExitCode = 2;
        }

        return true;
    }

    public static string GetStartupDiagnostics()
    {
        return _startupDiagnostics;
    }

    protected override void OnExit(ExitEventArgs e)
    {
        WriteSessionState("Clean", "Exit");
        base.OnExit(e);
    }

    public static void MarkOperation(string operation)
    {
        WriteSessionState("Running", operation);
        AppendActivityLog($"Operation: {operation}");
    }

    public static void MarkIdle()
    {
        WriteSessionState("Running", "Idle");
    }

    private static void WriteSessionState(string status, string operation)
    {
        try
        {
            EnsureReportFolder();
            var lines = new[]
            {
                $"Status={status}",
                $"Utc={DateTime.UtcNow:O}",
                $"Operation={operation}"
            };
            File.WriteAllLines(SessionStatePath, lines);
        }
        catch
        {
            // ignore state write failures
        }
    }

    private static (string Status, string Operation)? ReadSessionState()
    {
        try
        {
            if (!File.Exists(SessionStatePath))
                return null;

            var lines = File.ReadAllLines(SessionStatePath);
            var status = lines.FirstOrDefault(l => l.StartsWith("Status=", StringComparison.OrdinalIgnoreCase))?.Split('=', 2).ElementAtOrDefault(1) ?? "Unknown";
            var operation = lines.FirstOrDefault(l => l.StartsWith("Operation=", StringComparison.OrdinalIgnoreCase))?.Split('=', 2).ElementAtOrDefault(1) ?? "Unknown";
            return (status, operation);
        }
        catch
        {
            return null;
        }
    }

    private static void TryForceSoftwareRendering()
    {
        try
        {
            var property = typeof(RenderOptions).GetProperty("ProcessRenderMode", BindingFlags.Public | BindingFlags.Static);
            if (property?.PropertyType is null || !property.PropertyType.IsEnum)
            {
                AppendActivityLog("RenderMode switch not available.");
                return;
            }

            var softwareOnly = Enum.Parse(property.PropertyType, "SoftwareOnly", true);
            property.SetValue(null, softwareOnly);
            AppendActivityLog("RenderMode=SoftwareOnly");
        }
        catch (Exception ex)
        {
            AppendActivityLog($"RenderMode switch failed: {ex.Message}");
        }
    }

    public static string GetReportFolderPath()
    {
        EnsureReportFolder();
        return ReportFolder;
    }

    public static string SaveTextReport(string text)
    {
        EnsureReportFolder();
        var file = Path.Combine(ReportFolder, $"BendChecker-Report-{DateTime.Now:yyyyMMdd-HHmmss}.txt");
        File.WriteAllText(file, text ?? string.Empty);
        return file;
    }

    public static void AppendActivityLog(string text)
    {
        try
        {
            EnsureReportFolder();
            File.AppendAllText(ActivityLogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {text}{Environment.NewLine}");
        }
        catch
        {
            // ignore logging failures
        }
    }

    private static void EnsureReportFolder()
    {
        if (!Directory.Exists(ReportFolder))
            Directory.CreateDirectory(ReportFolder);
    }

    private static void RegisterGlobalExceptionHandlers()
    {
        Current.DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
    }

    private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        var text = e.Exception.ToString();
        AppendCrashLog("DispatcherUnhandledException", text);
        AppendActivityLog($"DispatcherUnhandledException: {e.Exception.Message}");
        ReportToVisualization("DispatcherUnhandledException", text);
        MessageBox.Show($"Unerwarteter Fehler. Log: {CrashLogPath}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }

    private static void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        var ex = e.ExceptionObject as Exception;
        var text = ex?.ToString() ?? "Unbekannter Fehler";
        AppendCrashLog("UnhandledException", text);
        AppendActivityLog($"UnhandledException: {ex?.Message ?? "Unbekannter Fehler"}");
        ReportToVisualization("UnhandledException", text);
        MessageBox.Show($"Unerwarteter Fehler. Log: {CrashLogPath}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        var text = e.Exception.ToString();
        AppendCrashLog("UnobservedTaskException", text);
        AppendActivityLog($"UnobservedTaskException: {e.Exception.Message}");
        ReportToVisualization("UnobservedTaskException", text);
        MessageBox.Show($"Task-Fehler. Log: {CrashLogPath}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        e.SetObserved();
    }

    private static void AppendCrashLog(string source, string details)
    {
        try
        {
            EnsureReportFolder();
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {source}{Environment.NewLine}{details}{Environment.NewLine}{new string('-', 80)}{Environment.NewLine}";
            File.AppendAllText(CrashLogPath, line);
        }
        catch
        {
            // ignore logging failures
        }
    }

    private static void ReportToVisualization(string source, string details)
    {
        if (Current?.Dispatcher is null)
            return;

        _ = Current.Dispatcher.BeginInvoke(() =>
        {
            if (Current.MainWindow is MainWindow window)
                window.ReportFromApp(source, details);
        });
    }

    private static void ConfigureOcctRuntime()
    {
        Debug.WriteLine($"ProcessPath: {Environment.ProcessPath}");
        Debug.WriteLine($"BaseDirectory: {AppContext.BaseDirectory}");

        var nativeDir = Path.Combine(AppContext.BaseDirectory, "runtimes", "win-x64", "native");
        if (Directory.Exists(nativeDir))
        {
            var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            var parts = path.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (!parts.Contains(nativeDir, StringComparer.OrdinalIgnoreCase))
                Environment.SetEnvironmentVariable("PATH", $"{nativeDir};{path}", EnvironmentVariableTarget.Process);
        }
        else
        {
            Debug.WriteLine($"Warning: Occt.NET runtime folder not found: {nativeDir}");
        }

        var arch = Environment.Is64BitProcess ? "x64" : "x86";
        var legacyNativeDir = Path.Combine(AppContext.BaseDirectory, "occt", arch);
        if (!Directory.Exists(legacyNativeDir))
        {
            Debug.WriteLine($"Warning: Occt.NET legacy folder not found: {legacyNativeDir}");
            return;
        }

        var legacyPath = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        var legacyParts = legacyPath.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (legacyParts.Contains(legacyNativeDir, StringComparer.OrdinalIgnoreCase))
            return;

        Environment.SetEnvironmentVariable("PATH", $"{legacyNativeDir};{legacyPath}", EnvironmentVariableTarget.Process);
    }
}
