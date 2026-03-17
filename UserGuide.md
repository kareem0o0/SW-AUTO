# User Guide

## What This Project Does

This project creates SOLIDWORKS files automatically.

It can currently generate:

- parts
- assemblies
- torsion-bar drawings

The system is based on simple objects.

The normal usage pattern is:

1. create the SOLIDWORKS application
2. create the shared PDM service
3. create a part or assembly object
4. change the properties you want
5. call the create method

## Project Layout

The main automation project is `SwAutomationApp/`.

Important locations:

- `SwAutomationApp/Parts/`
  - all part classes
- `SwAutomationApp/Assemblies/`
  - assembly helpers
- `SwAutomationApp/Drawing/`
  - drawing logic
- `SwAutomationApp/PDM/`
  - PDM logic
- `SwAutomationApp/macros.cs`
  - ready-to-edit workflow examples
- `SwAutomationApp/MachineAssembly.cs`
  - current application entry point

## Units

All linear values are now entered in meters.

Examples:

- `1.074` means 1074 mm
- `0.03` means 30 mm
- `0.99` means 990 mm

Angles are still entered in degrees.

This is very important when editing object properties.

## Basic Part Example

```csharp
SldWorks swApp = new SldWorks();
swApp.Visible = true;

PdmModule pdm = new PdmModule();

StatorSheetPart statorSheet = new StatorSheetPart(swApp, pdm);
statorSheet.OutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";
statorSheet.CloseAfterCreate = true;
statorSheet.OuterDiameter = 0.99;
statorSheet.InnerDiameter = 0.64;
statorSheet.PlateThickness = 0.1;

string partFile = statorSheet.Create();
```

## Basic Assembly Example

```csharp
AssemblyFile machine = new AssemblyFile(swApp, pdm);
machine.OutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";
machine.FileName = "MachineAssembly.SLDASM";
machine.CloseAfterCreate = false;

string assemblyFile = machine.Create();
```

After `Create()`, an open assembly can be used to:

- insert parts
- add mates
- create patterns

## Torsion Bar Special Behavior

`TorsionBarPart` currently has the most advanced workflow.

It supports two create options:

- `CreatePart()`
  - creates only the 3D part
- `Create()`
  - creates the 3D part and then creates the drawing

Example:

```csharp
TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
torsionBar.OutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";
torsionBar.DrawingOutputFolder = torsionBar.OutputFolder;
torsionBar.BarLength = 1.074;
torsionBar.BarHeight = 0.04;
torsionBar.BarThickness = 0.03;
torsionBar.DrawingLanguageCode = "EN";
torsionBar.DrawingCloseAfterCreate = false;

string drawingFile = torsionBar.Create();
```

## Torsion Bar Drawing Settings

The torsion-bar drawing settings live on the same `TorsionBarPart` object.

Common drawing properties:

- `DrawingOutputFolder`
- `DrawingSaveToPdm`
- `DrawingCloseAfterCreate`
- `DrawingLocalFileName`
- `DrawingSheetName`
- `DrawingLanguageCode`
- `DrawingTemplateFolderPath`
- `DrawingBottomTitleBlockClearance`
- `DrawingReferencedConfiguration`

This means you can configure both the part and the drawing from one object.

## PDM Data Card Values

Every part and assembly object owns a `PdmDataCard` property.

Example:

```csharp
statorSheet.PdmDataCard.Title = "Stator Sheet Metal";
statorSheet.PdmDataCard.Customer = "Birr Machines AG";
statorSheet.PdmDataCard.CustomerOrder = "66";
```

`TorsionBarPart` also has:

- `DrawingPdmDataCard`

The data-card object already contains default Birr values.

So if you do not set anything manually, the defaults stay in place.

## Important PDM Note

If you want to save to PDM:

1. create `PdmModule`
2. call `pdm.Login()`
3. set `SaveToPdm = true` on the object

If `SaveToPdm` is `false`, the file is saved locally only.

## Current Macros

The easiest way to use the project is often through the macros in `SwAutomationApp/macros.cs`.

### `Run4()`

`Run4()` is the main machine assembly workflow.

Use it when you want to:

- create the full machine part set
- build the machine assembly
- edit many object parameters in one place

`Run4()` creates the objects first, then creates the files, then inserts and mates the parts into the machine assembly.

### `Run5()`

`Run5()` is the torsion-bar drawing test workflow.

Use it when you want to:

- test the torsion-bar part
- test the torsion-bar drawing
- tune drawing settings such as language, template path, and close behavior

## Current Startup Behavior

The program starts from:

- `SwAutomationApp/MachineAssembly.cs`

That file currently calls:

- `Run4()`

If you want the program to launch the torsion-bar drawing workflow instead, change it to call:

- `Run5()`

## Current Part Classes

Available parts:

- `SkeletonPart`
- `StatorSheetPart`
- `ShaftPart`
- `StatorDistanceSheetPart`
- `StatorEndSheetPart`
- `TorsionBarPart`
- `PressPlatePart`
- `StatorPressringNdePart`

## Current Assembly Classes

Available assembly-related classes:

- `AssemblyFile`
- `RotorAssembly`
- `StatorAssembly`
- `HousingAssembly`

In most workflows, `AssemblyFile` is the main class used directly.

## Notes About the Macro Files

Only one explicit PDM data-card example is intentionally left in `Run4()`:

- the `SkeletonPart` block

That is there as a reference example.

If you later want to add PDM data-card overrides to another object, copy that same property-assignment pattern.

## Practical Advice

- if a value looks small, remember it is probably in meters now
- if you are testing only a drawing change, use `Run5()`
- if you are testing the full machine build, use `Run4()`
- if you save to PDM, make sure the vault login already happened
- if SolidWorks behaves differently on its first launch, try again after the application is fully open, because COM automation can be more sensitive on cold start

## Simple Mental Model

Use this as the quick mental model:

- part object = owns values and creates the part
- assembly object = creates assembly documents and helps with mates/patterns
- drawing logic = reads part data and creates the drawing
- macro = easy place to edit values and run a workflow

That is the current structure of the project.
