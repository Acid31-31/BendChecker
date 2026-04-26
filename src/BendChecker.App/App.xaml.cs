using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace BendChecker.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        RegisterGlobalExceptionHandlers();
        ConfigureOcctRuntime();
        base.OnStartup(e);
    }

    private static void RegisterGlobalExceptionHandlers()
    {
        Current.DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
    }

    private static void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        MessageBox.Show(e.Exception.ToString(), "Unerwarteter Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        e.Handled = true;
    }

    private static void OnUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        var ex = e.ExceptionObject as Exception;
        MessageBox.Show(ex?.ToString() ?? "Unbekannter Fehler", "Unerwarteter Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private static void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        MessageBox.Show(e.Exception.ToString(), "Task-Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        e.SetObserved();
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
