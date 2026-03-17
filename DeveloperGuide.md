# Developer Guide

## Purpose

This project generates SOLIDWORKS parts, assemblies, and a first drawing flow by using simple object-based classes.

The main rule is:

- each part is its own class
- each assembly is its own class
- editable values live on the object as properties
- the object reads its own properties when `Create()` is called

## Current File Layout

- `SwAutomationApp/Part.cs`
  - internal helpers
  - all part classes
  - the torsion-bar drawing entry point on `TorsionBarPart`
- `SwAutomationApp/Assembly.cs`
  - `AssemblyFile`
  - insert, mate, and pattern helpers
- `SwAutomationApp/drawing.cs`
  - drawing implementation methods
  - currently the torsion-bar drawing flow
- `SwAutomationApp/Pdm.cs`
  - `PdmModule`
  - `BirrDataCardValues`
  - PDM save and data-card update logic
- `SwAutomationApp/macros.cs`
  - editable orchestration flows
  - `Run4()` is the main machine assembly macro
  - `Run5()` is the torsion-bar drawing test macro
- `SwAutomationApp/Program.cs`
  - runtime entry point
  - currently calls `Run5()`

## Object Pattern

Most generator objects follow this shape:

1. Construct the object with `SldWorks` and `PdmModule`
2. Set editable properties
3. Call `Create()`

Example:

```csharp
StatorSheetPart statorSheet = new StatorSheetPart(swApp, pdm);
statorSheet.OutputFolder = outFolder;
statorSheet.PlateThicknessMm = 100.0;

string partPath = statorSheet.Create();
```

`AssemblyFile` follows the same pattern:

```csharp
AssemblyFile machine = new AssemblyFile(swApp, pdm);
machine.OutputFolder = outFolder;
machine.FileName = "MachineAssembly.SLDASM";

string assemblyPath = machine.Create();
```

## Part Layer

The part classes currently in `Part.cs` are:

- `SkeletonPart`
- `StatorSheetPart`
- `ShaftPart`
- `StatorDistanceSheetPart`
- `StatorEndSheetPart`
- `TorsionBarPart`
- `PressPlatePart`
- `StatorPressringNdePart`

Each part owns:

- file settings such as `OutputFolder`, `LocalFileName`, `CloseAfterCreate`, `SaveToPdm`
- geometry parameters
- `PdmDataCard`

`TorsionBarPart` also owns:

- drawing settings such as `DrawingOutputFolder`, `DrawingLanguageCode`, and drawing template settings
- `DrawingPdmDataCard`

`TorsionBarPart` has two public creation entry points:

- `CreatePart()`
  - creates only the 3D part
- `Create()`
  - creates the 3D part first
  - then calls the torsion-bar drawing logic in `drawing.cs`

## Assembly Layer

The assembly layer is currently one simple class:

- `AssemblyFile`

`AssemblyFile` owns:

- file settings such as `OutputFolder`, `FileName`, `CloseAfterCreate`, `SaveToPdm`
- `PdmDataCard`
- document creation
- component insertion
- mate helpers
- pattern helpers

## Drawing Layer

The drawing implementation stays in `SwAutomationApp/drawing.cs`.

Right now only the torsion-bar drawing is implemented.

Responsibility split:

- `TorsionBarPart`
  - owns the drawing-related data
  - exposes the public creation entry points
- `drawing.cs`
  - contains the actual drawing-generation logic

This keeps the part object as the data owner while keeping drawing implementation code out of `Part.cs`.

## PDM Data Card Pattern

`BirrDataCardValues` in `SwAutomationApp/Pdm.cs` stores the Birr data-card fields as object properties.

Defaults are already built into that class, so if a macro or external app does not set anything, the default values remain in place.

When `SaveToPdm` is `true`, the save path calls `PdmModule.SaveAsPdm(...)`, which also pushes the object's data-card values into PDM.

In `Run4()`, only `SkeletonPart` keeps an explicit `PdmDataCard` block in the macro. That block is there as the reference example if you want to copy the same pattern to another object later.

## Macros

### `Run4()`

`Run4()` is the main editable machine-build macro.

Flow:

1. Create all part and assembly objects
2. Edit their properties
3. Create the source files
4. Create the machine assembly
5. Insert components
6. Apply mates
7. Apply patterns

The assembly logic stays in the macro on purpose so it is easy to read and easy to edit.

### `Run5()`

`Run5()` is the torsion-bar drawing test macro.

Flow:

1. Create one `TorsionBarPart` object
2. Edit its 3D properties
3. Edit its drawing properties
4. Call `torsionBar.Create()`

## Extending the Code

### To add a new part

1. Add a new class in `Part.cs`
2. Add its editable properties
3. Add property defaults
4. Add `PdmDataCard`
5. Put the SOLIDWORKS creation logic inside `Create()`

### To add drawing support for a part

1. Keep the drawing-related properties on the part class
2. Add a clean entry method on the part if needed
3. Put the actual drawing implementation in `drawing.cs`

### To add a new macro flow

1. Create the needed objects in `macros.cs`
2. Set their properties directly in the macro
3. Call their creation methods
4. Keep the orchestration in the macro if that is the clearest place for it

## Important Notes

- Properties ending with `Mm` currently use millimeter values.
- Many SOLIDWORKS selections still depend on German default plane or reference names.
- If a workflow will save to PDM, `PdmModule.Login()` must be called before saving.
- There are still no automated geometry tests; validation is still done by running the automation and checking the created SOLIDWORKS files.
