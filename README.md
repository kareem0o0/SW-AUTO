# SwAutomationApp

Windows-only SOLIDWORKS automation project for generating parts and assemblies locally.

## Current Architecture

The project now uses a simple object-based model inside the existing source files.

- [Part.cs](c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/Part.cs)
  - small internal helpers
  - one concrete class per part
- [Assembly.cs](c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/Assembly.cs)
  - plain `AssemblyFile` class
- [macros.cs](c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/macros.cs)
  - editable orchestration flows
  - `Run4()` creates the machine assembly by creating part objects, then inserting and mating them
- [Program.cs](c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/Program.cs)
  - sample runtime entry

Each part or assembly is used like this:

1. Instantiate its class
2. Change properties if needed
3. Call `Create()`

Legacy property names still end with `Mm`, but the current input contract is meters.

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

## Example

```csharp
SldWorks swApp = new SldWorks();
swApp.Visible = true;
TorsionBarPart torsionBar = new TorsionBarPart(swApp);
torsionBar.OutputFolder = @"C:\temp\parts";
torsionBar.BarLengthMm = 1.074;

string torsionBarFile = torsionBar.Create();

AssemblyFile machine = new AssemblyFile(swApp);
machine.OutputFolder = @"C:\temp\parts";
machine.FileName = "MachineAssembly.SLDASM";

string machineFile = machine.Create();
```

For a full editable machine build flow, see `Run4()` in [macros.cs](c:/Users/kareem.salah/Downloads/birr%20machines/birr%20machines/SwAutomationApp/macros.cs).

## Run

```powershell
cd "C:\Users\kareem.salah\Downloads\birr machines\birr machines\SwAutomationApp"
dotnet run
```
