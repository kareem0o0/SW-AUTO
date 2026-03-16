# Developer Guide

## Purpose

This project uses class-based generators directly inside the existing source files.

## File Layout

- `SwAutomationApp/Part.cs`
  - small internal helper types
  - all concrete part classes and their `Create()` logic
- `SwAutomationApp/Assembly.cs`
  - `AssemblyFile`
  - assembly insert, mate, and pattern helpers
- `SwAutomationApp/macros.cs`
  - sample orchestration flows
  - `Run4()` contains the editable machine assembly sequence
- `SwAutomationApp/Program.cs`
  - creates `SldWorks`
  - creates `PdmModule`
  - runs a macro flow

## Main Rule

Each part or assembly is a class with:

- its own editable properties
- default values matching previous behavior
- its own `Create()` method

There is no central part-builder method anymore. Complex assembly orchestration can stay in `macros.cs` when that is easier to edit.

## Part Layer

The part classes in `Part.cs` are:

- `SkeletonPart`
- `StatorSheetPart`
- `ShaftPart`
- `StatorDistanceSheetPart`
- `StatorEndSheetPart`
- `TorsionBarPart`
- `PressPlatePart`
- `StatorPressringNdePart`

Each part class contains its own:

- `OutputFolder`
- `CloseAfterCreate`
- `SaveToPdm`
- `LocalFileName`
- UI suppression helpers

## Assembly Layer

The assembly class in `Assembly.cs` is:

- `AssemblyFile`

`AssemblyFile` contains its own:

- `OutputFolder`
- `CloseAfterCreate`
- `SaveToPdm`
- `FileName`
- assembly document creation
- component insertion
- mate helpers
- component pattern helpers

`Run4()` in `SwAutomationApp/macros.cs` is the main example of a generic assembly flow. It creates all part objects first, lets you edit any properties, then creates the assembly and applies inserts, mates, and patterns in one place.

## Extension Pattern

To add a new part:

1. Add a new class in `Part.cs`.
2. Add the editable properties it needs.
3. Add property defaults.
4. Move the SolidWorks creation logic into that class's `Create()` method.

To add a new assembly object:

1. Add a new class in `Assembly.cs`.
2. Add assembly-level properties such as `FileName`.
3. Put the assembly creation and helper operations inside that class.

To build a custom assembly flow:

1. Create the needed part objects in `macros.cs` or in the external app.
2. Change any parameters you need.
3. Create the parts.
4. Create an `AssemblyFile`.
5. Insert components and apply mates and patterns.

## Notes

- Editable values are exposed in mm.
- SOLIDWORKS API operations still use meters internally.
- Many selections depend on German SOLIDWORKS default names.
- There are still no automated geometry tests; validation is done by running SOLIDWORKS and inspecting the result.
