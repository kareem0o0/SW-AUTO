# SwAutomationApp

Windows-only SOLIDWORKS automation project for generating parts, assemblies, and drawings.

## Current Architecture

The project now uses a modular file layout with one main class per file.

- [Parts folder](/c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/Parts)
  - one file per part class
  - support helpers for shared automation utilities
- [Assemblies folder](/c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/Assemblies)
  - shared `AssemblyFile`
  - named assembly wrapper classes
- [PDM folder](/c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/PDM)
  - PDM save logic
  - data-card support
- [Drawing folder](/c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/Drawing)
  - drawing generation logic
- [macros.cs](/c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/macros.cs)
  - editable orchestration flows
- [MachineAssembly.cs](/c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/MachineAssembly.cs)
  - current executable entry point

## Part Classes

- `SkeletonPart`
- `StatorSheetPart`
- `ShaftPart`
- `StatorDistanceSheetPart`
- `StatorEndSheetPart`
- `TorsionBarPart`
- `PressPlatePart`
- `StatorPressringNdePart`

## Assembly Classes

- `AssemblyFile`
- `RotorAssembly`
- `StatorAssembly`
- `HousingAssembly`

## Example

```csharp
SldWorks swApp = new SldWorks();
swApp.Visible = true;

PdmModule pdm = new PdmModule();

TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
torsionBar.OutputFolder = @"C:\temp\parts";
torsionBar.BarLengthMm = 1074.0;

string torsionBarFile = torsionBar.CreatePart();

AssemblyFile machine = new AssemblyFile(swApp, pdm);
machine.OutputFolder = @"C:\temp\parts";
machine.FileName = "MachineAssembly.SLDASM";

string machineFile = machine.Create();
```

For a full editable machine build flow, see `Run4()` in [macros.cs](/c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/macros.cs).

## Run

```powershell
cd "C:\Users\kareem.salah\Downloads\birr machines\birr machines\SwAutomationApp"
dotnet run
```
