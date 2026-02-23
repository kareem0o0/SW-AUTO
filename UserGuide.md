# SOLIDWORKS Automation User Guide (Program.cs Flow)

## 1. How to use this project now
Use the project by editing `SwAutomationApp/Program.cs`.

Current pattern:
1. Create `PartParameters`
2. Call `partBuilder.GeneratePart(...)`
3. Build assembly using one of the assembly methods

---

## 2. Run
```powershell
cd "C:\Users\kareem.salah\Downloads\birr machines\birr machines\SwAutomationApp"
dotnet run
```

Output folder:
- `C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts`

---

## 3. Part generation pattern (new flow)
```csharp
// SketchPlaneName options:
// Front (default), Top, Right, Side (alias of Right)
PartParameters partAParameters = new PartParameters("Part_A", 100, 50, 10, 5, 3, SketchPlaneName.Front);
PartParameters partBParameters = new PartParameters("Part_B", 80, 80, 20, 4, 2, SketchPlaneName.Top);

string partAPath = partBuilder.GeneratePart(partAParameters, outFolder);
string partBPath = partBuilder.GeneratePart(partBParameters, outFolder);
```

Notes:
- If you omit `SketchPlaneName`, default is `Front`.
- Dimensions are in mm.

---

## 4. Assembly options in Program.cs

### Option A: Custom face pairs (recommended)
```csharp
// FaceName options:
// Top, Bottom, Left, Right, Front, Back, Up, Down
FaceMatePair[] customPairs =
{
    new FaceMatePair(FaceName.Top, FaceName.Bottom),
    new FaceMatePair(FaceName.Left, FaceName.Right, isFlipped: true)
};

assemblyBuilder.GenerateAssemblyByCustomFacePairs(
    partAPath,
    partBPath,
    outFolder,
    customPairs,
    "Final_Assembly_CustomFaceMates.SLDASM");
```

Flip option:
- `isFlipped: false` (default) -> aligned mate direction.
- `isFlipped: true` -> anti-aligned (flipped) mate direction.

Mixed selectors with flip:
```csharp
new FaceMatePair(
    new FaceSelection(50, 0, 5),
    new FaceSelection(FaceName.Right),
    isFlipped: true)
```

### Option B: Default face recipe
```csharp
assemblyBuilder.GenerateAssemblyByFaces(partAPath, partBPath, outFolder);
```

### Option C: Plane-based mating
```csharp
assemblyBuilder.GenerateAssembly(partAPath, partBPath, outFolder);
```

---

## 5. Mixed selector flow in custom pairs
`FaceMatePair` supports:
- Named face: `FaceName`
- Point-based face selection from part origin (mm): `FaceSelection(xMm, yMm, zMm)`

Example:
```csharp
FaceMatePair[] customPairs =
{
    new FaceMatePair(FaceName.Top, FaceName.Bottom),
    new FaceMatePair(new FaceSelection(50, 0, 5), new FaceSelection(FaceName.Right)),
    new FaceMatePair(new FaceSelection(50, 0, -5), new FaceSelection(40, 0, 10))
};

assemblyBuilder.GenerateAssemblyByCustomFacePairs(
    partAPath,
    partBPath,
    outFolder,
    customPairs,
    "Final_Assembly_CustomFaceMates.SLDASM");
```

Note:
- `FaceSelection(...)` values are mm from each part origin.
- Conversion to assembly coordinates is handled internally.

---

## 6. Sketch module (full step-by-step workflow)

Use `SketchBuilder` when you need explicit sketch control (entity type, relation type, and manual selection flow).

Create module:
```csharp
SketchBuilder sketchBuilder = new SketchBuilder(session.Application);
```

### 6.1 Start a new part sketch
```csharp
sketchBuilder.BeginPartSketch("SketchPart_1", outFolder, SketchPlaneName.Front);
```

Method:
```csharp
BeginPartSketch(string partName, string outputFolder, SketchPlaneName sketchPlane = SketchPlaneName.Front)
```

Arguments:
- `partName`: new part name without extension.
- `outputFolder`: target save folder.
- `sketchPlane`: `Front`, `Top`, `Right`, or `Side` (alias of `Right`).

### 6.2 Entity reference helper used by advanced sketch APIs
Several new methods require a typed entity selector:
```csharp
new SketchEntityReference("SKETCHSEGMENT", 0, -20)
new SketchEntityReference("SKETCHPOINT", 0, 0)
```

Method signature:
```csharp
SketchEntityReference(string type, double xMm, double yMm, double zMm = 0, string name = "")
```

Arguments:
- `type`: SOLIDWORKS selection type (examples: `SKETCHSEGMENT`, `SKETCHPOINT`).
- `xMm`, `yMm`, `zMm`: coordinate pick location in mm.
- `name`: optional named-entity selector (leave empty for coordinate-based picks).

### 6.3 Create sketch entities
Rectangles:
```csharp
sketchBuilder.CreateRectangle(SketchRectangleType.Center, 0, 0, 40, 20);
sketchBuilder.CreateRectangle(SketchRectangleType.Corner, -40, -20, 40, 20);
```

Circles:
```csharp
sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 8);
sketchBuilder.CreateCircle(SketchCircleType.CenterPoint, 0, 0, 8, 0);
```

### 6.4 Line method enhancements
Line types:
- `SketchLineType.Standard`
- `SketchLineType.Construction`
- `SketchLineType.Centerline`

Method:
```csharp
CreateLine(SketchLineType lineType, double startXmm, double startYmm, double endXmm, double endYmm)
```

Arguments:
- `lineType`: requested line behavior.
- `startXmm`, `startYmm`: line start in mm.
- `endXmm`, `endYmm`: line end in mm.

Examples:
```csharp
sketchBuilder.CreateLine(SketchLineType.Standard, -40, -20, 40, -20);
sketchBuilder.CreateLine(SketchLineType.Construction, -25, -20, -25, 20);
sketchBuilder.CreateLine(SketchLineType.Centerline, -40, 0, 40, 0);
```

Backward compatibility:
```csharp
sketchBuilder.CreateLine(-40, 0, 40, 0); // Standard line
```

### 6.5 Point method
Point types:
- `SketchPointType.Standard`
- `SketchPointType.Intersection`
- `SketchPointType.Midpoint`

Method:
```csharp
CreatePoint(
    SketchPointType pointType,
    double xMm,
    double yMm,
    SketchEntityReference firstReference = default,
    SketchEntityReference secondReference = default)
```

Arguments:
- `pointType`: point creation mode.
- `xMm`, `yMm`: point location in mm.
- `firstReference`: required for `Midpoint` and `Intersection`.
- `secondReference`: required for `Intersection`.

Examples:
```csharp
// Standard point.
sketchBuilder.CreatePoint(SketchPointType.Standard, 0, 0);

// Midpoint on one line.
sketchBuilder.CreatePoint(
    SketchPointType.Midpoint,
    0,
    -20,
    new SketchEntityReference("SKETCHSEGMENT", 0, -20));

// Intersection point constrained to two lines.
sketchBuilder.CreatePoint(
    SketchPointType.Intersection,
    0,
    0,
    new SketchEntityReference("SKETCHSEGMENT", 0, -20),
    new SketchEntityReference("SKETCHSEGMENT", 0, 20));
```

### 6.6 Fillet / Chamfer methods
Fillet method:
```csharp
CreateSketchFillet(
    double radiusMm,
    SketchEntityReference firstEntity,
    SketchEntityReference secondEntity,
    bool constrainedCorners = true)
```

Fillet example:
```csharp
sketchBuilder.CreateSketchFillet(
    5,
    new SketchEntityReference("SKETCHSEGMENT", -40, -20),
    new SketchEntityReference("SKETCHSEGMENT", -40, 20));
```

Chamfer modes:
- `SketchChamferMode.DistanceAngle`
- `SketchChamferMode.DistanceDistance`

Chamfer method:
```csharp
CreateSketchChamfer(
    SketchChamferMode chamferMode,
    double distance1Mm,
    double distance2OrAngle,
    SketchEntityReference firstEntity,
    SketchEntityReference secondEntity)
```

Chamfer arguments:
- `distance1Mm`: first distance in mm.
- `distance2OrAngle`:
  - if `DistanceDistance`: second distance in mm.
  - if `DistanceAngle`: angle in degrees.

Chamfer examples:
```csharp
sketchBuilder.CreateSketchChamfer(
    SketchChamferMode.DistanceAngle,
    4,
    45,
    new SketchEntityReference("SKETCHSEGMENT", 40, -20),
    new SketchEntityReference("SKETCHSEGMENT", 40, 20));

sketchBuilder.CreateSketchChamfer(
    SketchChamferMode.DistanceDistance,
    4,
    2,
    new SketchEntityReference("SKETCHSEGMENT", 40, -20),
    new SketchEntityReference("SKETCHSEGMENT", 40, 20));
```

### 6.7 Relation / constraint helper method
Supported options:
- `Horizontal`
- `Vertical`
- `Parallel`
- `Perpendicular`
- `Tangent`
- `Equal`
- `Symmetric`
- `Coincident`
- `Collinear`

Method:
```csharp
ApplySketchRelation(SketchRelationType relationType, params SketchEntityReference[] entities)
```

Examples:
```csharp
sketchBuilder.ApplySketchRelation(
    SketchRelationType.Parallel,
    new SketchEntityReference("SKETCHSEGMENT", 0, -20),
    new SketchEntityReference("SKETCHSEGMENT", 0, 20));

sketchBuilder.ApplySketchRelation(
    SketchRelationType.Coincident,
    new SketchEntityReference("SKETCHPOINT", 0, 0),
    new SketchEntityReference("SKETCHSEGMENT", 0, -20));
```

Legacy low-level method still available:
```csharp
sketchBuilder.AddSketchConstraint("sgCOINCIDENT");
```

### 6.8 Selection and dimensions
Selection methods:
```csharp
SelectInSketch(string name, string type, double xMm, double yMm, double zMm, bool append, int mark)
SelectSketchPoint(double xMm, double yMm, bool append = false, int mark = 0)
```

Dimension method:
```csharp
AddDimension(double textXmm, double textYmm, double textZmm = 0)
```

### 6.9 End sketch, create feature, save
```csharp
sketchBuilder.EndSketch();
sketchBuilder.Extrude(12, midPlane: false);
string savedPath = sketchBuilder.SaveAndClose(closeDocument: false);
```

### 6.10 Complete workflow example
```csharp
sketchBuilder.BeginPartSketch("SketchPart_1", outFolder, SketchPlaneName.Front);
sketchBuilder.CreateRectangle(SketchRectangleType.Center, 0, 0, 40, 20);
sketchBuilder.CreateLine(SketchLineType.Centerline, -40, 0, 40, 0);
sketchBuilder.CreateLine(SketchLineType.Standard, -40, -20, 40, -20);
sketchBuilder.CreateLine(SketchLineType.Standard, -40, 20, 40, 20);
sketchBuilder.CreatePoint(SketchPointType.Standard, 0, 0);
sketchBuilder.CreatePoint(
    SketchPointType.Midpoint,
    0,
    -20,
    new SketchEntityReference("SKETCHSEGMENT", 0, -20));
sketchBuilder.ApplySketchRelation(
    SketchRelationType.Parallel,
    new SketchEntityReference("SKETCHSEGMENT", 0, -20),
    new SketchEntityReference("SKETCHSEGMENT", 0, 20));
sketchBuilder.CreateSketchFillet(
    5,
    new SketchEntityReference("SKETCHSEGMENT", -40, -20),
    new SketchEntityReference("SKETCHSEGMENT", -40, 20));
sketchBuilder.CreateSketchChamfer(
    SketchChamferMode.DistanceAngle,
    4,
    45,
    new SketchEntityReference("SKETCHSEGMENT", 40, -20),
    new SketchEntityReference("SKETCHSEGMENT", 40, 20));
sketchBuilder.AddDimension(0, 25, 0);
sketchBuilder.EndSketch();
sketchBuilder.Extrude(12, midPlane: false);
sketchBuilder.SaveAndClose(closeDocument: false);
```

---

## 7. Other Program.cs creation flows

### Centered parts only
```csharp
partBuilder.CreateCenteredRectangularPart("RectPart_Centered", 120, 60, 20, outFolder);
partBuilder.CreateCenteredCircularPart("CircPart_Centered", 50, 100, outFolder);
```

### New part with two offset planes
```csharp
partBuilder.CreatePartWithOffsetPlanes(0.05); // 50 mm
```

### Add offset plane to existing part
```csharp
string rectPartPath = Path.Combine(outFolder, "RectPart_Centered.SLDPRT");
partBuilder.CreateReferenceOffsetPlane(rectPartPath, "Front", 0.10);
```

---

## 8. Recommended editing workflow
1. Edit only `Program.cs` scenario block.
2. Keep one assembly flow active at a time.
3. Run `dotnet run`.
4. Check generated files in `parts` folder.

---

## 9. Common errors

### Compile error in `AssemblyBuilder.cs`: `Invalid expression term '{'`
Ensure ray arrays are written as:
- `rays = new[] { ... }`

### SDK missing
```powershell
dotnet --list-sdks
```

Install .NET SDK if list is empty.
