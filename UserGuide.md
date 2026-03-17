# SOLIDWORKS Automation User Guide

## Basic Usage

The project works with simple objects.

The normal pattern is:

1. Create `SldWorks`
2. Create `PdmModule`
3. Create a part or assembly object
4. Change any properties you want
5. Call the object's creation method

If you want to save files to PDM, log into PDM first with `pdm.Login()`.

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

## Torsion Bar Special Case

`TorsionBarPart` currently has two public creation options:

- `CreatePart()`
  - creates only the 3D part
- `Create()`
  - creates the 3D part
  - then creates the drawing too

Example:

```csharp
TorsionBarPart torsionBar = new TorsionBarPart(swApp, pdm);
torsionBar.OutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";
torsionBar.DrawingOutputFolder = torsionBar.OutputFolder;
torsionBar.BarLengthMm = 1074.0;
torsionBar.DrawingLanguageCode = "EN";

string drawingFile = torsionBar.Create();
```

## PDM Data Card Values

Every part and assembly object now owns a `PdmDataCard` property.

Example:

```csharp
statorSheet.PdmDataCard.Title = "Stator Sheet Metal";
statorSheet.PdmDataCard.Customer = "Birr Machines AG";
statorSheet.PdmDataCard.CustomerOrder = "66";
```

`TorsionBarPart` also has `DrawingPdmDataCard` for the drawing file.

The built-in defaults remain in place if you do not override them.

For readability, only one explicit macro example is left in `Run4()` on `SkeletonPart`. You can copy that same pattern to any other object later if needed.

## Macros

### `Run4()`

Use `Run4()` when you want the full editable machine assembly flow.

It:

1. Creates all part objects
2. Lets you edit their parameters in one place
3. Creates the source files
4. Creates the machine assembly
5. Inserts components
6. Adds mates and patterns

### `Run5()`

Use `Run5()` when you want to test the torsion-bar drawing flow.

It:

1. Creates one `TorsionBarPart`
2. Lets you edit both the part and drawing settings
3. Calls `Create()`

## Current Runtime Entry

`SwAutomationApp/Program.cs` currently starts `Run5()`.

If you want to test the machine assembly flow instead, change `Program.cs` to call `Run4(...)`.

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

## Important Notes

- Properties ending with `Mm` currently use millimeter values.
- If `SaveToPdm` is `false`, the file is saved locally only.
- If `SaveToPdm` is `true`, make sure PDM login is already done.
