# Developer Guide

## Purpose

This project automates the creation of SOLIDWORKS parts, assemblies, and drawings.

The codebase is organized around simple objects:

- each part is its own class in its own file
- each assembly helper is isolated in its own file
- object properties hold the editable values
- `Create()` reads the current object state and generates the file

The design goal is to keep the code easy to extend from two directions:

- a developer can add new geometry or workflows without editing one giant file
- an external Windows application can create an object, set properties, and call the create method

## Current Project Structure

The main automation project lives in `SwAutomationApp/`.

Important folders and files:

- `SwAutomationApp/Parts/`
  - one file per part class
  - shared helpers such as `AutomationSupport.cs` and `AutomationUiScope.cs`
- `SwAutomationApp/Assemblies/`
  - `AssemblyFile.cs` for generic assembly document behavior
  - wrapper classes such as `RotorAssembly.cs`, `StatorAssembly.cs`, and `HousingAssembly.cs`
- `SwAutomationApp/PDM/`
  - `Pdm.cs` for PDM login, save, and data-card updates
- `SwAutomationApp/Drawing/`
  - `drawing.cs` for drawing-generation logic
- `SwAutomationApp/macros.cs`
  - editable orchestration flows
- `SwAutomationApp/MachineAssembly.cs`
  - executable entry point

## Main Architecture

### 1. Part objects own their data

Each part class stores the values needed to build that part.

Typical property groups on a part object:

- file settings
  - `OutputFolder`
  - `LocalFileName`
  - `SaveToPdm`
  - `CloseAfterCreate`
- geometry settings
  - diameters
  - offsets
  - thickness values
  - pattern counts
  - material name
- PDM data-card settings
  - `PdmDataCard`

This means the part object is the single source of truth for that part.

### 2. Drawing logic stays outside the part file

Drawing logic is implemented in `SwAutomationApp/Drawing/drawing.cs`.

However, the drawing-related values still live on the owning part object.

The best example is `TorsionBarPart`:

- the part owns all 3D parameters
- the same part also owns all drawing settings
- `drawing.cs` reads those values and creates the drawing

This keeps responsibilities clean:

- part class = owns data
- drawing file = performs drawing operations

### 3. Assemblies are document helpers, not workflow owners

`AssemblyFile` is a reusable helper for assembly documents.

It knows how to:

- create an assembly document
- save locally or to PDM
- insert parts or subassemblies
- create mates
- create linear and circular patterns

It does not decide the full machine structure.

That orchestration stays in `macros.cs`, especially in `Run4()`.

## Units

This project now uses meters directly for all linear values.

Important rule:

- all lengths, diameters, offsets, spacing values, and thicknesses are stored in meters
- angles are still stored in degrees where the API expects angle input from the user side

Examples:

- `0.99` = 990 mm
- `0.04` = 40 mm
- `1.074` = 1074 mm

This is an important update.

Older code and older documentation used millimeter-based property names and conversion helpers.
That is no longer the current architecture.

There is no `MmToMeters` helper in the current code path.

## Entry Point

The executable entry point is:

- `SwAutomationApp/MachineAssembly.cs`

It is intentionally small.

Its job is:

1. create the SOLIDWORKS application object
2. create the shared `PdmModule`
3. choose which macro to run

At the moment it starts:

- `Project1.Run4(...)`

If you want the executable to launch the torsion-bar drawing test instead, change that one call to:

- `Project1.Run5(...)`

## Parts

Current part classes:

- `SkeletonPart`
- `StatorSheetPart`
- `ShaftPart`
- `StatorDistanceSheetPart`
- `StatorEndSheetPart`
- `TorsionBarPart`
- `PressPlatePart`
- `StatorPressringNdePart`

Each file contains one class only.

This makes it much easier to:

- find a specific part
- add properties to one part without affecting others
- debug one geometry generator at a time

## Assemblies

Current assembly files:

- `AssemblyFile`
- `RotorAssembly`
- `StatorAssembly`
- `HousingAssembly`

Important note:

- `AssemblyFile` is the real shared implementation
- the wrapper assembly classes currently exist mainly to give named identities and clearer structure

The actual full machine build is still assembled in `Run4()`.

## PDM Layer

The PDM logic is in:

- `SwAutomationApp/PDM/Pdm.cs`

This file contains two important types.

### `BirrDataCardValues`

This is a plain object that stores the Birr PDM data-card fields as properties.

Current fields include:

- `DrawingNumber`
- `Title`
- `Subtitle`
- `Project`
- `Customer`
- `CustomerOrder`
- `Type`
- `Unit`
- `CreatedFrom`
- `ReplacementFor`
- `DataCheck`

These properties already have default values.

That means:

- if a macro does nothing, the defaults remain
- if an external system sets values, those override the defaults

### `PdmModule`

This class handles:

- vault login
- saving a live SOLIDWORKS document into PDM
- reading data-card values
- writing data-card values

When `SaveToPdm` is `true`, the object's `PdmDataCard` values are pushed into the PDM file.

## PDM Pattern in Macros

Only one explicit PDM data-card assignment block is intentionally left in `Run4()`:

- the `SkeletonPart` block

That block acts as the example/reference pattern.

If you want to apply the same idea to another object later, copy that same property assignment style to that object.

This was done on purpose to keep the macro readable instead of repeating the same data-card block on every object.

## Drawing Layer

The drawing implementation is currently focused on the torsion bar.

File:

- `SwAutomationApp/Drawing/drawing.cs`

Current idea:

- `TorsionBarPart.CreatePart()` creates only the 3D part
- `TorsionBarPart.Create()` creates the part and then calls the drawing logic

The drawing logic currently:

1. ensures the part exists
2. selects a sheet size and scale
3. resolves the drawing template and sheet format
4. creates the drawing
5. inserts the front view
6. creates projected views from that base view
7. saves the drawing locally or to PDM

The code also supports:

- English and German sheet formats through `DrawingLanguageCode`
- template folder configuration
- drawing-specific output settings

## `TorsionBarPart` Special Case

`TorsionBarPart` is currently the most advanced example in the project.

It owns:

- part geometry properties
- part save properties
- drawing properties
- part PDM data card
- drawing PDM data card

Its public entry points are:

- `CreatePart()`
  - creates only the part
- `Create()`
  - creates the part and then the drawing

This is the cleanest current example of the project's intended architecture.

## Macros

Macros are kept in:

- `SwAutomationApp/macros.cs`

They are not low-level geometry files.
They are workflow files.

Their purpose is to show how the part and assembly objects can be used together.

### `Run4()`

`Run4()` is the main machine-assembly flow.

It is intentionally written as a readable orchestration script.

High-level flow:

1. create all part objects
2. assign their editable properties
3. create the source files
4. create the target machine assembly
5. insert the components
6. apply mates
7. create patterns

This macro is a good place for:

- changing machine dimensions
- changing stack counts
- changing material names
- copying the PDM data-card pattern

### `Run5()`

`Run5()` is the torsion-bar drawing test flow.

It creates one `TorsionBarPart`, sets both part and drawing properties, and then calls:

- `torsionBar.Create()`

Use this flow when working on:

- torsion-bar geometry
- template selection
- drawing layout
- drawing save behavior

## Example Usage Pattern

The preferred style in this project is explicit property assignment.

Example:

```csharp
TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
torsionBar.OutputFolder = outFolder;
torsionBar.SaveToPdm = false;
torsionBar.CloseAfterCreate = true;
torsionBar.BarLength = 1.074;
torsionBar.BarHeight = 0.04;
torsionBar.BarThickness = 0.03;

string partPath = torsionBar.CreatePart();
```

For a drawing flow:

```csharp
TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
torsionBar.OutputFolder = outFolder;
torsionBar.DrawingOutputFolder = outFolder;
torsionBar.DrawingLanguageCode = "EN";
torsionBar.DrawingCloseAfterCreate = false;

string drawingPath = torsionBar.Create();
```

## Extending the Code

### Add a new part

1. Create a new file under `SwAutomationApp/Parts/`
2. Add one class only in that file
3. Add the editable properties
4. Give the properties sensible defaults
5. Add `PdmDataCard`
6. Implement the SOLIDWORKS geometry creation inside `Create()`

### Add drawing support to a part

1. keep drawing-related settings on the part object
2. keep detailed drawing logic in `SwAutomationApp/Drawing/drawing.cs`
3. add or reuse a clean entry point on the part class

### Add a new assembly wrapper

1. add a file under `SwAutomationApp/Assemblies/`
2. reuse `AssemblyFile` where possible
3. keep large orchestration logic in macros unless there is a strong reason to move it

## Common Development Notes

- many reference selections still rely on SolidWorks plane or axis names such as German default names
- projected drawing views work more reliably when the source model stays open
- PDM save flows require `pdm.Login()` before saving
- SolidWorks automation can behave differently on a fresh first launch, so some creation flows include stability steps such as explicit document activation and rebuilds

## Mental Model

If you are new to the project, this is the easiest way to think about it:

- part class = data owner + 3D generator
- drawing file = drawing engine
- assembly file = assembly document helper
- macro = workflow/orchestration layer
- machine entry file = startup switchboard

That mental model matches the current codebase and should be followed for future additions.
