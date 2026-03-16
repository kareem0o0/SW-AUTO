# SOLIDWORKS Automation User Guide

## Usage Pattern

1. Create `SldWorks`
2. Create `PdmModule`
3. Create a part or assembly object
4. Change any properties you want
5. Call `Create()`

## Example Part

```csharp
SldWorks swApp = new SldWorks();
swApp.Visible = true;
PdmModule pdm = new PdmModule();

StatorSheetPart statorSheet = new StatorSheetPart(swApp, pdm);
statorSheet.OutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";
statorSheet.PlateThicknessMm = 100.0;

string statorSheetFile = statorSheet.Create();
```

## Example Assembly

```csharp
AssemblyFile machine = new AssemblyFile(swApp, pdm);
machine.OutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";
machine.FileName = "MachineAssembly.SLDASM";

string machineFile = machine.Create();
```

For a full editable machine build flow with part objects, parameter overrides, insertions, mates, and patterns in one place, use `Run4()` in `SwAutomationApp/macros.cs`.

## Available Classes

Parts:

- `SkeletonPart`
- `StatorSheetPart`
- `ShaftPart`
- `StatorDistanceSheetPart`
- `StatorEndSheetPart`
- `TorsionBarPart`
- `PressPlatePart`
- `StatorPressringNdePart`

Assemblies:

- `AssemblyFile`
