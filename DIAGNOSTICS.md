# Diagnostics Guide

## Start der App
- Visual Studio: `BendChecker.App` als Startprojekt, dann **F5**.
- CLI: `dotnet run --project src/BendChecker.App/BendChecker.App.csproj`

## Wo liegt die Diagnose-Datei?
Beim Start werden zwei Dateien geschrieben:
- `%TEMP%\BendChecker_diagnostics.txt`
- `AppContext.BaseDirectory\diagnostics.txt` (z. B. `src/BendChecker.App/bin/Debug/net8.0-windows/diagnostics.txt`)

## Welche Infos enthält diagnostics.txt?
- Timestamp
- `Environment.ProcessPath`
- `AppContext.BaseDirectory`
- `.NET Environment.Version`
- OS-Version
- `Environment.Is64BitProcess`
- Ob `PATH` den Ordner `runtimes\win-x64\native` enthält
- Existenz von:
  - `runtimes\win-x64\native`
  - `occt\x64`
- Dateiliste (Name + Größe) aus `runtimes\win-x64\native`
- NativeLibrary-Probe-Load (erste 5 DLLs): `LOAD OK` / `LOAD FAIL` + Exception

## Wenn Crash/Faulting Module auftritt
### Faulting module = `VCRUNTIME140.dll` / `MSVCP140.dll`
- Microsoft Visual C++ Redistributable 2015-2022 (x64) installieren/reparieren.

### Faulting module = OCCT/Qt/TK DLL
- Prüfen, ob `runtimes\win-x64\native` vollständig im Output vorhanden ist.
- Prüfen, ob die DLL im Probe-Load als `LOAD FAIL` auftaucht.
- Wenn ja: Abhängigkeit fehlt oder falsche Architektur (x86/x64 mismatch).

### Faulting module = `KERNELBASE.dll` mit InnerException DllNotFound
- In `diagnostics.txt` den Abschnitt `NativeLibrary Probe` prüfen.
- Fehlende DLLs oder PATH-/Deployment-Problem beheben.

## Nächster Schritt
- `diagnostics.txt` + ggf. EventViewer-Fehler (Faulting module + Exception code) anhängen.
- Danach gezielt fehlende Runtime-Komponente oder DLL-Abhängigkeit beheben.
