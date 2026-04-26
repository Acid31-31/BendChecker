using System.IO;
using System.Windows;

namespace BendChecker.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        ConfigureOcctRuntime();
        base.OnStartup(e);
    }

    private static void ConfigureOcctRuntime()
    {
        var arch = Environment.Is64BitProcess ? "x64" : "x86";
        var nativeDir = Path.Combine(AppContext.BaseDirectory, "occt", arch);
        if (!Directory.Exists(nativeDir))
            return;

        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        var parts = path.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Contains(nativeDir, StringComparer.OrdinalIgnoreCase))
            return;

        Environment.SetEnvironmentVariable("PATH", $"{nativeDir};{path}", EnvironmentVariableTarget.Process);
    }
}

