// Other methods and code

public void ProbeNativeDllLoad()
{
    // Check the environment variable
    if (Environment.GetEnvironmentVariable("BENDCHECKER_DISABLE_NATIVE_PROBE") == "1")
    {
        return; // Skip the probe
    }

    string[] dllsToLoad = {
        "TKernel.dll",
        "TKMath.dll",
        "TKXSBase.dll",
        "TKDESTEP.dll",
        "TxOcct.dll"
    };

    foreach (var dll in dllsToLoad)
    {
        string dllPath = Path.Combine("runtimes/win-x64/native", dll);
        try
        {
            if (File.Exists(dllPath))
            {
                // Load the DLL
                LoadLibrary(dllPath);
                AppendDiagnostics($"{dll}: LOAD OK");
            }
            else
            {
                AppendDiagnostics($"{dll}: MISSING");
            }
        }
        catch (Exception ex)
        {
            AppendDiagnostics($"{dll}: LOAD FAIL\n{ex.ToString()}");
        }
    }

    // Other methods and code
}