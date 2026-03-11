# SwAutomationApp

Windows-only SOLIDWORKS automation project for generating parts and assemblies locally or in SOLIDWORKS PDM.

## Current Status

The project currently supports:

- starting a visible SOLIDWORKS session from code
- creating parts directly from `Part.cs`
- creating assemblies directly from `Assembly.cs`
- saving either locally or to SOLIDWORKS PDM through `Pdm.cs`
- switching runnable scenarios through `macros.cs`

## Implemented Part Generators

File: `SwAutomationApp/Part.cs`

- `Create_stator_sheet(...)`
  - Original stator ring with slot geometry and the ear-style extension logic
  - Circular pattern of the slot cut
- `Create_shaft(...)`
  - Stepped shaft built as 5 consecutive cylindrical sections
- `CreateSkeleton(...)`
  - Reference-plane-only skeleton part for downstream assembly mating
- `Create_stator_distance_sheet(...)`
  - Stator distance sheet ring
  - Rectangular slot with circular pattern
  - Extra rectangular boss body
  - Boss-side cut profile
  - Circular body pattern
- `Create_stator_end_sheet(...)`
  - End sheet version of the stator ring
  - Slot only, no extra boss body
- `Create_torsion_bar(...)`
  - Rectangular torsion bar
  - Two configurations:
    - `P0001`: plain bar
    - `P0002`: center plain hole plus threaded holes created with Hole Wizard

## Implemented Assembly Helpers

File: `SwAutomationApp/Assembly.cs`

- `CreateAssembly(...)`
- `InsertComponentToOpenAssembly(...)`
- `mate_plans(...)`
- `ApplyCoincedentMate(...)`
- `ApplyParallelMate(...)`
- `ApplyCoincedentAxisMate(...)`

These helpers support simple assembly creation, inserting generated parts/assemblies, and mating by named planes, axes, or named faces.

## PDM Support

File: `SwAutomationApp/Pdm.cs`

- `Login()`
- `SaveAsPdm(...)`
- `AddExistingFileToPdm(...)`
- `GetDataCardValues(...)`
- `UpdateBirrDataCard(...)`
- `FillBirrDataCard(...)`

The current implementation contains hardcoded vault credentials and a hardcoded local vault root path.

## Current Runtime Entry

File: `SwAutomationApp/Program.cs`

The project currently runs:

```csharp
Project1.Run3(localoutFolder, myPart, assembly);
```

File: `SwAutomationApp/macros.cs`

Available scenario entry points:

- `Run(...)`
  - skeleton + multiple assemblies + PDM-oriented flow
- `Run2(...)`
  - skeleton + stator sheet + shaft + simple mating flow
- `Run3(...)`
  - currently used for the torsion bar flow

## Current Project Layout

```text
SwAutomationApp/
  Program.cs
  macros.cs
  Part.cs
  Assembly.cs
  Pdm.cs
  SwAutomationApp.csproj
README.md
UserGuide.md
DeveloperGuide.md
```

## Requirements

- Windows
- installed SOLIDWORKS
- installed SOLIDWORKS PDM client if using PDM flows
- local .NET SDK compatible with the project target framework
- COM interop assemblies available at the paths referenced in the project

## Run

```powershell
cd "C:\Users\kareem.salah\Downloads\birr machines\birr machines\SwAutomationApp"
dotnet run
```

## Notes

- Many selections rely on German default SOLIDWORKS names such as `Ebene vorne`, `Ebene oben`, `Ebene rechts`, and `Z-Achse`.
- Most geometry methods define their editable dimensions at the top of each method.
- `UserGuide.md` contains older user-focused notes and examples.
- `DeveloperGuide.md` is the current technical reference for maintaining and extending the codebase.
