# SOLIDWORKS Automation API Guide

## Table of Contents
1. [Overview](#1-overview)
2. [Requirements](#2-requirements)
3. [Quick Start](#3-quick-start)
4. [Project Structure](#4-project-structure)
5. [Core Data Models](#5-core-data-models)
6. [Session Lifecycle](#6-session-lifecycle)
7. [PartBuilder API](#7-partbuilder-api)
8. [AssemblyBuilder API](#8-assemblybuilder-api)
9. [SketchBuilder API](#9-sketchbuilder-api)
10. [Examples Runner Methods](#10-examples-runner-methods)
11. [Program.cs Usage Pattern](#11-programcs-usage-pattern)
12. [Troubleshooting](#12-troubleshooting)

## 1. Overview
This guide documents the current implementation in:
- `SwAutomationApp/Program.cs`
- `SwAutomationApp/Examples.cs`
- `SwAutomationApp/PartBuilder.cs`
- `SwAutomationApp/AssemblyBuilder.cs`
- `SwAutomationApp/SketchBuilder.cs`
- `SwAutomationApp/Models.cs`
- `SwAutomationApp/SwSession.cs`

## 2. Requirements
1. Windows OS
2. SOLIDWORKS installed and COM-accessible
3. .NET SDK compatible with the project target framework (`net10.0-windows`)

## 3. Quick Start
```powershell
cd "C:\Users\kareem.salah\Downloads\birr machines\birr machines\SwAutomationApp"
dotnet run
```

Default output folder in examples:
- `C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts`

## 4. Project Structure
- `Program.cs`: app entry and scenario switching
- `Examples.cs`: reusable scenario methods
- `PartBuilder.cs`: part operations
- `AssemblyBuilder.cs`: assembly operations
- `SketchBuilder.cs`: sketch-level operations
- `Models.cs`: shared enums/data objects
- `SwSession.cs`: SOLIDWORKS COM lifecycle management

## 5. Core Data Models
### 5.1 `FaceName`
Values: `Top`, `Bottom`, `Left`, `Right`, `Front`, `Back`, `Up`, `Down`

### 5.2 `SketchPlaneName`
Values: `Front`, `Top`, `Right`, `Side` (`Side` aliases to `Right`)

### 5.3 `PartParameters`
```csharp
PartParameters(string name, double widthMm, double depthMm, double heightMm, double holeDiameterMm, int holeCount, SketchPlaneName sketchPlane = SketchPlaneName.Front)
```
- `name`: required, part filename without extension
- `widthMm/depthMm/heightMm`: required, mm
- `holeDiameterMm`: required by constructor, mm (`0` disables holes)
- `holeCount`: required by constructor, count (`0` disables holes)
- `sketchPlane`: optional, default `Front`

### 5.4 `FaceSelection`
```csharp
FaceSelection(FaceName faceName)
FaceSelection(double xMm, double yMm, double zMm)
```
- face mode or point mode (mm from part origin)

### 5.5 `FaceMatePair`
```csharp
FaceMatePair(FaceName partAFace, FaceName partBFace)
FaceMatePair(FaceName partAFace, FaceName partBFace, bool isFlipped)
FaceMatePair(FaceSelection partAFace, FaceSelection partBFace, bool isFlipped = false)
```

## 6. Session Lifecycle
```csharp
using SwSession session = new SwSession(visible: true);
```
- Connects to `SldWorks.Application`
- Releases COM resources on dispose

## 7. PartBuilder API

### 7.1 `GenerateRectangularPartWithHoles`
```csharp
string GenerateRectangularPartWithHoles(PartParameters parameters, string folder)
```
- `parameters`: required
- `folder`: required output directory
- Returns saved part path (or empty string on skip/failure)
- This is the primary explicit method name for rectangular block + holes generation.
- Backward-compatible alias still exists:
```csharp
string GeneratePart(PartParameters parameters, string folder)
```

### 7.2 `GeneratePart` (start new blank part + sketch)
```csharp
ModelDoc2 GeneratePart(string partName, SketchPlaneName sketchPlane = SketchPlaneName.Front)
```
- `partName`: required new part name
- `sketchPlane`: optional base sketch plane
- Returns the new active `ModelDoc2` part document with sketch mode started

### 7.3 `ExtrudeExistingSketchAndSave`
```csharp
string ExtrudeExistingSketchAndSave(ModelDoc2 partModel, string partName, string outputFolder, double depthMm, bool midPlane = false)
```
- `partModel`: required part document containing existing sketch geometry
- `partName`: required output part name
- `outputFolder`: required save directory
- `depthMm`: required extrusion depth in mm
- `midPlane`: optional end condition
- Returns saved `.SLDPRT` full path

Example:
```csharp
ModelDoc2 started = partBuilder.GeneratePart("StartedPart_1", SketchPlaneName.Front);
started.SketchManager.CreateCenterRectangle(0, 0, 0, 0.02, 0.01, 0); // 20x10 mm (meters in raw COM)
string saved = partBuilder.ExtrudeExistingSketchAndSave(started, "StartedPart_1", outFolder, 12, midPlane: false);
```

### 7.4 `CreatePartWithOffsetPlanes`
```csharp
void CreatePartWithOffsetPlanes(double offsetDistance)
```
- `offsetDistance`: required, meters

### 7.5 `CreateReferenceOffsetPlane`
```csharp
void CreateReferenceOffsetPlane(string componentPath, string planeName, double offsetDistance)
```
- `componentPath`: required `.SLDPRT` path
- `planeName`: required
- `offsetDistance`: required, meters

### 7.6 `CreateCenteredRectangularPart`
```csharp
void CreateCenteredRectangularPart(string partName, double x, double y, double z, string outputFolder)
```
- `x/y/z`: required, mm

### 7.7 `CreateCenteredCircularPart`
```csharp
void CreateCenteredCircularPart(string partName, double diameter, double z, string outputFolder)
```
- `diameter/z`: required, mm

### 7.8 `MirrorPartFeature` (new)
```csharp
void MirrorPartFeature(string partPath, SketchEntityReference mirrorPlaneReference, SketchEntityReference seedFeatureReference, bool geometryPattern = true, bool merge = true, bool knit = false, int scopeOptions = 0)
```
- `partPath`: required `.SLDPRT` path
- `mirrorPlaneReference`: required mirror plane selector (`PLANE`)
- `seedFeatureReference`: required feature/body selector (`BODYFEATURE`, etc.)
- `geometryPattern/merge/knit/scopeOptions`: optional mirror options

Example:
```csharp
partBuilder.MirrorPartFeature(
    rectPartPath,
    new SketchEntityReference("PLANE", 0, 0, 0, "Front Plane"),
    new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"));
```

### 7.9 `CreatePartLinearPattern` (new)
```csharp
void CreatePartLinearPattern(string partPath, SketchEntityReference seedFeatureReference, SketchEntityReference direction1Reference, int count1, double spacing1Mm, SketchEntityReference direction2Reference = default, int count2 = 1, double spacing2Mm = 0, bool flipDir1 = false, bool flipDir2 = false)
```
- `partPath`: required `.SLDPRT` path
- `seedFeatureReference`: required seed feature/body
- `direction1Reference`: required first direction selector
- `count1`: required, >= 1
- `spacing1Mm`: required, mm
- `direction2Reference`: optional second direction selector
- `count2`: optional, default `1`
- `spacing2Mm`: optional, mm
- `flipDir1/flipDir2`: optional direction flip flags

Example:
```csharp
partBuilder.CreatePartLinearPattern(
    rectPartPath,
    new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"),
    new SketchEntityReference("EDGE", 0, 0, 0),
    4,
    20);
```

### 7.10 `CreatePartCircularPattern` (new)
```csharp
void CreatePartCircularPattern(string partPath, SketchEntityReference seedFeatureReference, SketchEntityReference axisReference, int count, double spacingDeg, bool flipDir = false)
```
- `partPath`: required `.SLDPRT` path
- `seedFeatureReference`: required seed feature/body
- `axisReference`: required axis selector
- `count`: required, >= 1
- `spacingDeg`: required, degrees
- `flipDir`: optional

Example:
```csharp
partBuilder.CreatePartCircularPattern(
    rectPartPath,
    new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"),
    new SketchEntityReference("AXIS", 0, 0, 0),
    6,
    60);
```

## 8. AssemblyBuilder API

### 8.1 `GenerateAssembly`
```csharp
void GenerateAssembly(string partAPath, string partBPath, string folder)
```

### 8.2 `GenerateAssemblyByFaces`
```csharp
void GenerateAssemblyByFaces(string partAPath, string partBPath, string folder)
```

### 8.3 `GenerateAssemblyByCustomFacePairs`
```csharp
void GenerateAssemblyByCustomFacePairs(string partAPath, string partBPath, string folder, FaceMatePair[] matePairs, string outputAssemblyFileName)
```

Example:
```csharp
FaceMatePair[] pairs =
{
    new FaceMatePair(FaceName.Top, FaceName.Bottom),
    new FaceMatePair(new FaceSelection(50, 0, 5), new FaceSelection(FaceName.Right), isFlipped: true)
};

assemblyBuilder.GenerateAssemblyByCustomFacePairs(partAPath, partBPath, outFolder, pairs, "Custom.SLDASM");
```

## 9. SketchBuilder API

### 9.1 Enums
- `SketchRectangleType`: `Center`, `Corner`
- `SketchCircleType`: `CenterRadius`, `CenterPoint`
- `SketchLineType`: `Standard`, `Construction`, `Centerline`
- `SketchChamferMode`: `DistanceAngle`, `DistanceDistance`
- `SketchRelationType`: `Horizontal`, `Vertical`, `Parallel`, `Perpendicular`, `Tangent`, `Equal`, `Symmetric`, `Coincident`, `Collinear`

### 9.2 `SketchEntityReference`
```csharp
SketchEntityReference(double xMm, double yMm, double zMm = 0, string name = "")
SketchEntityReference(string type, double xMm, double yMm, double zMm = 0, string name = "")
```
- Defaults `type` to `SKETCHPOINT` when omitted/empty

### 9.3 Core sketch lifecycle
```csharp
void BeginPartSketch(string partName, string outputFolder, SketchPlaneName sketchPlane = SketchPlaneName.Front)
void EndSketch()
void Extrude(double depthMm, bool midPlane = false, bool isCut = false)
string SavePart()
string SaveAndClose(bool closeDocument)
ModelDoc2 GetActiveModel()
void Rebuild()
```

Extrude arguments:
- `depthMm`: required, extrusion depth in mm
- `midPlane`: optional, default `false`
- `isCut`: optional, default `false`
  - `false` => boss/base extrude
  - `true` => cut extrude

`Extrude(...)` behavior note:
- If the extruded result is non-cylindrical (for example rectangular), helper reference planes are auto-created:
  - `Face_Front`, `Face_Back`, `Face_Right`, `Face_Left`, `Face_Top`, `Face_Bottom`
- If the extruded result contains cylindrical surfaces, only these helper planes are created:
  - `Face_Top`, `Face_Bottom`
- This makes follow-up sketch plane selection easier.

### 9.4 Geometry
```csharp
void CreateRectangle(SketchRectangleType rectangleType, double x1Mm, double y1Mm, double x2Mm, double y2Mm)
void CreateCircle(SketchCircleType circleType, double centerXmm, double centerYmm, double value1Mm, double value2Mm = 0)
void CreateLine(SketchLineType lineType, double startXmm, double startYmm, double endXmm, double endYmm)
void CreateLine(double startXmm, double startYmm, double endXmm, double endYmm)
void CreatePoint(double xMm, double yMm)
```

### 9.5 Mirror/pattern (new)

#### 9.5.1 `MirrorSketchEntities`
```csharp
void MirrorSketchEntities(SketchEntityReference mirrorEntity, params SketchEntityReference[] entitiesToMirror)
```
- `mirrorEntity`: required (usually centerline/segment)
- `entitiesToMirror`: required, one or more

Example:
```csharp
sketchBuilder.MirrorSketchEntities(
    new SketchEntityReference("SKETCHSEGMENT", 0, 0),
    new SketchEntityReference("SKETCHSEGMENT", 60, -20));
```

#### 9.5.2 `CreateLinearSketchPattern`
```csharp
void CreateLinearSketchPattern(int numX, int numY, double spacingXmm, double spacingYmm, double angleXdeg = 0, double angleYdeg = 90, string deleteInstances = "", params SketchEntityReference[] seedEntities)
```
- `numX/numY`: required counts (>= 1)
- `spacingXmm/spacingYmm`: required, mm
- `angleXdeg/angleYdeg`: optional, degrees
- `deleteInstances`: optional SOLIDWORKS delete-instance string
- `seedEntities`: required, one or more

Example:
```csharp
sketchBuilder.CreateLinearSketchPattern(
    3, 1, 20, 0, 0, 90, "",
    new SketchEntityReference("SKETCHSEGMENT", 60, -20));
```

#### 9.5.3 `CreateCircularSketchPattern`
```csharp
void CreateCircularSketchPattern(double arcRadiusMm, double arcAngleDeg, int patternCount, double patternSpacingDeg, bool patternRotate = false, string deleteInstances = "", params SketchEntityReference[] seedEntities)
```
- `arcRadiusMm`: required, mm
- `arcAngleDeg`: required, degrees
- `patternCount`: required count (>= 1)
- `patternSpacingDeg`: required, degrees
- `patternRotate`: optional
- `deleteInstances`: optional
- `seedEntities`: required, one or more

Example:
```csharp
sketchBuilder.CreateCircularSketchPattern(
    40, 180, 6, 30, false, "",
    new SketchEntityReference("SKETCHSEGMENT", 60, -20));
```

### 9.6 Relations and constraints
```csharp
void ApplySketchRelation(SketchRelationType relationType, params SketchEntityReference[] entities)
void AddSketchConstraint(string constraintToken)
```

### 9.7 Smart Dimensions

#### 9.7.1 Preselected references
```csharp
void AddDimension(double textXmm, double textYmm, double textZmm = 0)
```

#### 9.7.2 Two-reference distance dimension
```csharp
void AddDimension(double firstXmm, double firstYmm, double secondXmm, double secondYmm, double textXmm, double textYmm, double textZmm = 0, string firstType = "", string secondType = "")
```

#### 9.7.3 Single-entity dimension
```csharp
void AddDimension(double entityXmm, double entityYmm, double textXmm, double textYmm, double textZmm = 0, string entityType = "")
```
- supports line length and arc/circle radius/diameter smart dimensions

Examples:
```csharp
sketchBuilder.AddDimension(-40, 20, 40, 20, 0, 30, 0);
sketchBuilder.AddDimension(0, 20, 0, 35, 0);
```

### 9.8 Selection helpers
```csharp
bool SelectInSketch(string name, string type, double xMm, double yMm, double zMm, bool append, int mark)
bool SelectSketchPoint(double xMm, double yMm, bool append = false, int mark = 0)
```

## 10. Examples Runner Methods
Available in `Examples.cs`:
- `RunCustomFacePairsCurrent()`
- `RunAssemblyByPlanes()`
- `RunAssemblyByDefaultFaceMates()`
- `RunAssemblyByCustomFaceMappingVariant()`
- `RunCenteredStockPartsOnly()`
- `RunCreatePartWithOffsetPlanes()`
- `RunAddOffsetPlanesToExistingPart()`
- `RunSketchWorkflow()`

## 11. Program.cs Usage Pattern
1. Open `SwAutomationApp/Program.cs`
2. Activate one scenario at a time
3. Run `dotnet run`
4. Inspect output in parts folder

## 12. Troubleshooting
1. Mirror/pattern fails:
- verify reference types (`PLANE`, `BODYFEATURE`, `EDGE`, `AXIS`, `SKETCHSEGMENT`)
- verify pick coordinates are close to intended geometry

2. Smart dimension fails:
- use two-reference overload for distance dimensions
- use single-entity overload for line/arc/circle dimensions
- pass explicit type when inference is ambiguous

3. SDK not found:
```powershell
dotnet --list-sdks
```

4. SOLIDWORKS connection failure:
- verify installation and COM registration of `SldWorks.Application`
