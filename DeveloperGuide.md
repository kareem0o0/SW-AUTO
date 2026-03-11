# Developer Guide

## Purpose

This guide explains the current structure of the SOLIDWORKS automation project, how the implemented generators work, and what a developer should know before modifying or extending the code.

## 1. Repository Overview

Main code lives in `SwAutomationApp/`.

- `Program.cs`
  - application entry point
  - creates the SOLIDWORKS session
  - creates `PdmModule`, `Part`, and `Assembly`
  - chooses which scenario macro to run
- `macros.cs`
  - scenario switchboard
  - contains `Run`, `Run2`, and `Run3`
  - this is the easiest place to change what `dotnet run` actually generates
- `Part.cs`
  - all part-generation logic
  - currently the most important file in the project
- `Assembly.cs`
  - assembly creation and mating helpers
- `Pdm.cs`
  - PDM login, save, and data-card helpers

## 2. Current Runtime Flow

At the time of writing, `Program.cs` creates the app services and then calls:

```csharp
Project1.Run3(localoutFolder, myPart, assembly);
```

`Run3(...)` currently calls:

```csharp
part.Create_torsion_bar(outFolder, false, SaveToPdm: false);
```

If you want to run a different generator or assembly flow, change `Program.cs` or switch the active call inside `macros.cs`.

## 3. Common Patterns Used In Part.cs

Most part methods follow the same structure:

1. Define a local `Mm(...)` converter from mm to meters.
2. Use `BeginAutomationUiSuppression()` to reduce SOLIDWORKS UI interference during automation.
3. Put all editable dimensions at the top of the method.
4. Convert all driving values to meters in a derived-values block.
5. Create a new part from the default SOLIDWORKS part template.
6. Build sketches and features step by step.
7. Save either locally or through `_pdm.SaveAsPdm(...)`.

Important implementation details:

- SOLIDWORKS API dimensions are in meters.
- Many methods use a local `SelectSketchByIndex(...)` helper because sketch names may be either German (`Skizze1`) or English (`Sketch1`).
- Many selections are coordinate-based, especially faces, sketch segments, and sketch points.
- Several methods temporarily disable sketch inference:
  - `_swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);`
- The code often restores the view at the end with:
  - `ShowNamedView2("*Front", ...)`
  - `ViewZoomtofit2()`

## 4. Implemented Part Generators

### 4.1 `Create_stator_sheet(...)`

Purpose:
- Creates the original stator ring sheet.

Key behavior:
- sketches inner and outer circles
- extrudes the ring
- creates the slot geometry
- includes the original ear-style extension logic from the source drawing
- cuts the slot and circular-patterns it around the ring

Notes:
- this is the reference method that several later stator methods were derived from
- it still contains more complex slot/ear geometry than the distance and end sheets

### 4.2 `Create_shaft(...)`

Purpose:
- Builds a stepped shaft from multiple cylindrical sections.

Key behavior:
- uses 5 editable radii
- uses 5 editable section lengths
- creates one circular sketch per section
- extrudes each section in sequence using the previous end face

Notes:
- this method is straightforward and useful as a reference for serial feature creation

### 4.3 `CreateSkeleton(...)`

Purpose:
- Creates a reference-only skeleton part used for assembly mating.

Key behavior:
- creates reference planes offset from default planes
- names them:
  - `NDE_BEARING_CENTER`
  - `DE_BEARING_CENTER`
  - `Ground_Plane`

Notes:
- this is not a solid modeling part
- it is mainly used as an assembly reference object

### 4.4 `Create_stator_distance_sheet(...)`

Purpose:
- Creates the stator distance sheet.

Key behavior:
- builds the ring from inner and outer circles
- creates the slot geometry without the original ear extension
- keeps the slot sketch stable by converting the bottom rectangle edge to construction instead of trimming it away
- circular-patterns the slot cut
- creates an extra rectangular boss as a separate body
- adds a side-face cut profile to the boss
- circular-patterns the boss body

Important implementation details:

- the slot sketch now preserves its defining dimensions by turning the bottom edge into construction
- the boss body is intentionally created as a separate body
- the boss cut was mirrored to match the corrected drawing orientation
- the final boss pattern is a body pattern, not just a feature pattern

### 4.5 `Create_stator_end_sheet(...)`

Purpose:
- Creates the stator end sheet.

Key behavior:
- same ring-and-slot family as the distance sheet
- no extra boss body
- slot sketch uses the same construction-bottom-edge approach as the distance sheet
- slot cut is circular-patterned around the ring

Important implementation details:

- the slot definition is intentionally kept similar to the distance-sheet slot logic
- this method is the lighter-weight version of the distance sheet

### 4.6 `Create_torsion_bar(...)`

Purpose:
- Creates a torsion bar with two configurations.

Configurations:
- `P0001`
  - plain bar
- `P0002`
  - same bar plus holes from the drawing

Current geometry assumptions:
- base bar size is `1074 x 40 x 30` mm
- center hole is a plain `16 mm` through cut
- `2 x M10` are modeled as tapped hole-wizard holes
- `2 x M16` are modeled as tapped hole-wizard holes

Current implementation strategy:

1. Rename the default active configuration to `P0001`.
2. Build the rectangular bar.
3. Create the center plain hole as a normal cut.
4. Create the threaded holes using `FeatureManager.HoleWizard5(...)`.
5. Add configuration `P0002`.
6. Suppress hole features in `P0001`.
7. Unsuppress the same features in `P0002`.

Important implementation details:

- threaded holes use Hole Wizard, not plain extruded cuts
- size strings currently use:
  - `M10x1.5` with fallback `M10`
  - `M16x2` with fallback `M16`
- configuration-specific feature states are controlled with `Feature.SetSuppression2(...)`

## 5. Assembly Layer

File: `Assembly.cs`

Main responsibilities:

- create a new assembly and optionally save it to PDM
- insert generated parts or assemblies into the currently open assembly
- apply simple mating operations

Key methods:

- `CreateAssembly(...)`
  - creates a new assembly document
  - saves it locally or to PDM
  - optionally keeps it open for further insert/mate operations
- `InsertComponentToOpenAssembly(...)`
  - opens the target part or assembly silently
  - inserts it into the active assembly
  - unfixed the component afterward
- `mate_plans(...)`
  - mates default assembly planes to the inserted component planes
- `ApplyCoincedentMate(...)`
- `ApplyParallelMate(...)`
- `ApplyCoincedentAxisMate(...)`

Important implementation detail:

- mating logic supports both German plane names and named faces/entities

## 6. PDM Layer

File: `Pdm.cs`

Responsibilities:

- login to the vault
- save SOLIDWORKS documents into the local vault view
- register files with PDM
- inspect and update data-card values

Key methods:

- `Login()`
- `SaveAsPdm(...)`
- `AddExistingFileToPdm(...)`
- `GetDataCardValues(...)`
- `UpdateBirrDataCard(...)`
- `FillBirrDataCard(...)`

Important implementation details:

- credentials are currently hardcoded
- the local vault root path is currently hardcoded
- `UpdateBirrDataCard(...)` currently updates `@` and `P0001`

## 7. How To Add A New Part Generator

Recommended pattern:

1. Add a new public method in `Part.cs`.
2. Keep all editable dimensions grouped at the top.
3. Convert mm to meters immediately in a derived-values block.
4. Reuse the existing save pattern:
   - local save with `SaveAs3(...)`
   - PDM save with `_pdm.SaveAsPdm(...)`
5. If the feature tree depends on sketch order, use a local `SelectSketchByIndex(...)`.
6. If SOLIDWORKS auto-relations cause trouble, temporarily disable sketch inference.
7. If the part has variants, prefer configuration-specific suppression over duplicating the whole method.
8. Expose the new method through `macros.cs` or `Program.cs` for testing.

## 8. Common SolidWorks Automation Pitfalls In This Project

### 8.1 German entity names

Many selections depend on default German names:

- `Ebene vorne`
- `Ebene oben`
- `Ebene rechts`
- `Z-Achse`

If SOLIDWORKS runs with different default naming, those selections may fail.

### 8.2 Coordinate-based selection is fragile

A lot of the code selects:

- faces by a point on the face
- sketch segments by a point on the segment
- sketch points by exact coordinates

If the geometry changes significantly, those selectors may need to change too.

### 8.3 Sketch dimensions can be lost when entities are trimmed away

This already happened in the stator slot work.

The current solution in the distance-sheet and end-sheet slot sketches is:

- keep the bottom rectangle edge
- convert it to construction
- preserve the dimensions and relations that were attached to it

### 8.4 Feature names are not always stable

Calls like:

- `Cut-Extrude1`
- `Sketch2`
- `Boss-Extrude1`

assume the feature order stays stable.

If you insert new features in the middle of a method, downstream named selections may need to be updated.

### 8.5 Live environment required

There are no automated tests for feature creation.

Validation usually means:

1. `dotnet run`
2. watch SOLIDWORKS build the model
3. inspect the resulting feature tree, bodies, configurations, and final file

## 9. Suggested Workflow For Developers

When changing or adding geometry:

1. Edit one generator at a time.
2. Keep parameters grouped and clearly named.
3. Run only the relevant scenario from `macros.cs`.
4. Test with local save first.
5. Switch to PDM save only after the geometry is stable.
6. Update the README and this guide when adding new generators or workflows.

## 10. Known Documentation Notes

- `UserGuide.md` contains older notes from an earlier abstraction layer and is not the main source of truth for the current direct-generator architecture.
- This file is intended to be the technical reference for the current codebase state.
