// Import base .NET types.
using System;
// Import file helpers.
using System.IO;
// Import SOLIDWORKS interop types.
using SolidWorks.Interop.sldworks;
// Import SOLIDWORKS constant enums.
using SolidWorks.Interop.swconst;

namespace SwAutomation;

// Rectangle creation variants supported by SketchBuilder.
public enum SketchRectangleType
{
    Center,
    Corner
}

// Circle creation variants supported by SketchBuilder.
public enum SketchCircleType
{
    CenterRadius,
    CenterPoint
}

// Line creation variants supported by SketchBuilder.
public enum SketchLineType
{
    Standard,
    Construction,
    Centerline
}

// Sketch chamfer creation variants.
public enum SketchChamferMode
{
    DistanceAngle,
    DistanceDistance
}

// Supported sketch relation options.
public enum SketchRelationType
{
    Horizontal,
    Vertical,
    Parallel,
    Perpendicular,
    Tangent,
    Equal,
    Symmetric,
    Coincident,
    Collinear
}

// Reference to select an existing sketch entity.
public readonly struct SketchEntityReference
{
    public const string DefaultType = "SKETCHPOINT";

    public string Name { get; }

    public string Type { get; }

    public double XMm { get; }

    public double YMm { get; }

    public double ZMm { get; }

    public SketchEntityReference(string type, double xMm, double yMm, double zMm = 0, string name = "")
    {
        // If type is omitted/empty, default to sketch point selection.
        Type = string.IsNullOrWhiteSpace(type) ? DefaultType : type;
        XMm = xMm;
        YMm = yMm;
        ZMm = zMm;
        Name = name ?? string.Empty;
    }

    // Convenience overload: lets callers omit the entity type (defaults to SKETCHPOINT).
    public SketchEntityReference(double xMm, double yMm, double zMm = 0, string name = "")
        : this(DefaultType, xMm, yMm, zMm, name)
    {
    }

    public bool IsValid()
    {
        return !string.IsNullOrWhiteSpace(Type);
    }
}

// Step-by-step sketch module for building complex sketches before feature creation.
public sealed class SketchBuilder
{
    // Shared SOLIDWORKS app instance.
    private readonly SldWorks _swApp;

    // Active part document for current sketch workflow.
    private ModelDoc2 _activeModel;

    // Current part name for save/close operations.
    private string _activePartName = string.Empty;

    // Current output folder for save operations.
    private string _outputFolder = string.Empty;

    // Tracks whether sketch mode is currently active.
    private bool _isSketchOpen;

    // Create sketch module with SOLIDWORKS dependency.
    public SketchBuilder(SldWorks swApp)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
    }

    // Create a new part, select requested reference plane, and begin sketch mode.
    public void BeginPartSketch(string partName, string outputFolder, SketchPlaneName sketchPlane = SketchPlaneName.Front)
    {
        if (string.IsNullOrWhiteSpace(partName)) throw new ArgumentException("Part name is required.", nameof(partName));
        if (string.IsNullOrWhiteSpace(outputFolder)) throw new ArgumentException("Output folder is required.", nameof(outputFolder));

        // Ensure output folder exists.
        if (!Directory.Exists(outputFolder)) Directory.CreateDirectory(outputFolder);

        // Create new part from default template.
        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        _activeModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
        _activePartName = partName;
        _outputFolder = outputFolder;

        // Select requested base plane.
        string[] preferredPlaneNames = GetSketchPlaneCandidates(sketchPlane);
        bool planeSelected = SelectPartPlane(_activeModel, preferredPlaneNames, false);
        if (!planeSelected)
        {
            // Fallback to first available reference plane.
            planeSelected = SelectFirstReferencePlane(_activeModel, false);
            if (!planeSelected)
            {
                throw new InvalidOperationException("Failed to select sketch base plane.");
            }
        }

        // Enter sketch mode.
        _activeModel.SketchManager.InsertSketch(true);
        _isSketchOpen = true;
        Console.WriteLine("Sketch started on " + sketchPlane + " plane.");
    }

    // Begin a new sketch on an existing part using one of the standard datum planes.
    // If partModel is not provided, the current active model in this builder is used.
    // Example:
    //   sketchBuilder.BeginSketch(SketchPlaneName.Top, existingPartModel);
    //   sketchBuilder.BeginSketch(SketchPlaneName.Front);
    public void BeginSketch(SketchPlaneName sketchPlane = SketchPlaneName.Front, ModelDoc2 partModel = null)
    {
        if (partModel != null)
        {
            _activeModel = partModel;
        }

        EnsureModelReady();

        if (_isSketchOpen)
        {
            EndSketch();
        }

        _activeModel.ClearSelection2(true);

        string[] preferredPlaneNames = GetSketchPlaneCandidates(sketchPlane);
        bool selected = SelectPartPlane(_activeModel, preferredPlaneNames, false);
        if (!selected)
        {
            selected = SelectFirstReferencePlane(_activeModel, false);
            if (!selected)
            {
                throw new InvalidOperationException("Failed to select sketch plane on existing part.");
            }
        }

        _activeModel.SketchManager.InsertSketch(true);
        _isSketchOpen = true;
        Console.WriteLine("Sketch started on existing part (" + sketchPlane + ").");
    }

    // Begin a new sketch on an existing part using an explicit reference (plane or planar face).
    // Example for a plane:
    //   sketchBuilder.BeginSketch(new SketchEntityReference("PLANE", 0, 0, 0, name: "Top Plane"), existingPartModel);
    // Example for a face pick by coordinate:
    //   sketchBuilder.BeginSketch(new SketchEntityReference("FACE", 40, 0, 10), existingPartModel);
    public void BeginSketch(SketchEntityReference sketchTargetReference, ModelDoc2 partModel = null)
    {
        if (partModel != null)
        {
            _activeModel = partModel;
        }

        EnsureModelReady();

        if (!sketchTargetReference.IsValid())
        {
            throw new ArgumentException("Sketch target reference is required.", nameof(sketchTargetReference));
        }

        if (_isSketchOpen)
        {
            EndSketch();
        }

        _activeModel.ClearSelection2(true);

        bool isPlaneTarget = string.Equals(sketchTargetReference.Type, "PLANE", StringComparison.OrdinalIgnoreCase);
        bool selected = isPlaneTarget
            ? SelectPlaneTarget(sketchTargetReference)
            : SelectInSketch(
                sketchTargetReference.Name,
                sketchTargetReference.Type,
                sketchTargetReference.XMm,
                sketchTargetReference.YMm,
                sketchTargetReference.ZMm,
                false,
                0);

        if (!selected)
        {
            string hint = isPlaneTarget
                ? " Try SketchPlaneName.Top/Front/Right overload or use a valid planar FACE reference."
                : string.Empty;
            throw new InvalidOperationException("Failed to select sketch target reference (" + sketchTargetReference.Type + ")." + hint);
        }

        // Rectangle/circle style 2D sketch entities require a planar sketch target.
        // If user selected FACE, validate that it is planar to avoid silent command stalls.
        if (string.Equals(sketchTargetReference.Type, "FACE", StringComparison.OrdinalIgnoreCase))
        {
            SelectionMgr selMgr = _activeModel.SelectionManager as SelectionMgr;
            Face2 selectedFace = selMgr?.GetSelectedObject6(1, -1) as Face2;
            Surface surface = selectedFace?.GetSurface() as Surface;
            if (surface == null || !surface.IsPlane())
            {
                _activeModel.ClearSelection2(true);
                throw new InvalidOperationException(
                    "BeginSketch target FACE is not planar. Select a planar face interior point or use Face_Top/Face_Bottom plane.");
            }
        }

        _activeModel.SketchManager.InsertSketch(true);
        _isSketchOpen = true;
        Console.WriteLine("Sketch started on selected target (" + sketchTargetReference.Type + ").");
    }

    // Robust plane-target selection for BeginSketch:
    // 1) explicit name + PLANE
    // 2) ref-plane feature exact/prefix
    // 3) Face_* alias to standard datum planes
    // 4) first reference plane fallback
    private bool SelectPlaneTarget(SketchEntityReference planeReference)
    {
        // First try direct SelectByID2 with name/type.
        bool selected = SelectInSketch(
            planeReference.Name,
            planeReference.Type,
            planeReference.XMm,
            planeReference.YMm,
            planeReference.ZMm,
            false,
            0);
        if (selected) return true;

        // Try tree-based exact/prefix match (handles names like Face_Top<1>).
        if (!string.IsNullOrWhiteSpace(planeReference.Name))
        {
            selected = SelectReferencePlaneByNameOrPrefix(planeReference.Name, false);
            if (selected) return true;

            // Face_* aliases map to base datum planes when helper planes do not exist.
            string[] aliasCandidates = GetFacePlaneAliasCandidates(planeReference.Name);
            if (aliasCandidates.Length > 0)
            {
                selected = SelectPartPlane(_activeModel, aliasCandidates, false);
                if (selected) return true;
            }
        }

        // Last resort: pick first available reference plane to avoid hard-stop.
        return SelectFirstReferencePlane(_activeModel, false);
    }

    // Create rectangle entity using selected rectangle type.
    // Coordinates are in mm in sketch space.
    public void CreateRectangle(SketchRectangleType rectangleType, double x1Mm, double y1Mm, double x2Mm, double y2Mm)
    {
        EnsureSketchOpen();

        // Convert mm to meters for SOLIDWORKS API.
        double x1 = x1Mm / 1000.0;
        double y1 = y1Mm / 1000.0;
        double x2 = x2Mm / 1000.0;
        double y2 = y2Mm / 1000.0;

        int beforeCount = GetActiveSketchSegmentCount();

        // Build rectangle by requested variant.
        object rectangleResult;
        switch (rectangleType)
        {
            case SketchRectangleType.Center:
                rectangleResult = _activeModel.SketchManager.CreateCenterRectangle(x1, y1, 0, x2, y2, 0);
                break;
            case SketchRectangleType.Corner:
            default:
                rectangleResult = _activeModel.SketchManager.CreateCornerRectangle(x1, y1, 0, x2, y2, 0);
                break;
        }

        int afterCount = GetActiveSketchSegmentCount();

        // SOLIDWORKS can intermittently fail rectangle API calls on some face-based sketches.
        // Fallback path: build the rectangle from 4 lines when API returns null
        // or when no new segments were committed.
        if (rectangleResult == null || (beforeCount >= 0 && afterCount <= beforeCount))
        {
            CreateRectangleByLines(rectangleType, x1Mm, y1Mm, x2Mm, y2Mm);
        }
    }

    // Create circle entity using selected circle type.
    // Coordinates and radius are in mm in sketch space.
    public void CreateCircle(SketchCircleType circleType, double centerXmm, double centerYmm, double value1Mm, double value2Mm = 0)
    {
        EnsureSketchOpen();

        double cx = centerXmm / 1000.0;
        double cy = centerYmm / 1000.0;
        double v1 = value1Mm / 1000.0;
        double v2 = value2Mm / 1000.0;

        switch (circleType)
        {
            case SketchCircleType.CenterRadius:
                _activeModel.SketchManager.CreateCircleByRadius(cx, cy, 0, v1);
                break;
            case SketchCircleType.CenterPoint:
            default:
                _activeModel.SketchManager.CreateCircle(cx, cy, 0, v1, v2, 0);
                break;
        }
    }

    // Create line entity in sketch space (mm) using requested line mode.
    public void CreateLine(SketchLineType lineType, double startXmm, double startYmm, double endXmm, double endYmm)
    {
        EnsureSketchOpen();

        double startX = startXmm / 1000.0;
        double startY = startYmm / 1000.0;
        double endX = endXmm / 1000.0;
        double endY = endYmm / 1000.0;

        switch (lineType)
        {
            case SketchLineType.Centerline:
                _activeModel.SketchManager.CreateCenterLine(startX, startY, 0, endX, endY, 0);
                break;

            case SketchLineType.Construction:
                _activeModel.SketchManager.CreateLine(startX, startY, 0, endX, endY, 0);
                _activeModel.ClearSelection2(true);
                SelectInSketch(string.Empty, "SKETCHSEGMENT", (startXmm + endXmm) / 2.0, (startYmm + endYmm) / 2.0, 0, false, 0);
                _activeModel.SketchManager.CreateConstructionGeometry();
                _activeModel.ClearSelection2(true);
                break;

            case SketchLineType.Standard:
            default:
                _activeModel.SketchManager.CreateLine(startX, startY, 0, endX, endY, 0);
                break;
        }
    }

    // Backward-compatible line method (standard line).
    public void CreateLine(double startXmm, double startYmm, double endXmm, double endYmm)
    {
        CreateLine(SketchLineType.Standard, startXmm, startYmm, endXmm, endYmm);
    }

    // Create sketch point in mm coordinates.
    public void CreatePoint(double xMm, double yMm)
    {
        EnsureSketchOpen();
        _activeModel.SketchManager.CreatePoint(xMm / 1000.0, yMm / 1000.0, 0);
    }

    // Mirror selected sketch entities about a selected mirror entity (usually a centerline).
    // Example:
    //   MirrorSketchEntities(
    //       new SketchEntityReference("SKETCHSEGMENT", 0, 0),
    //       new SketchEntityReference("SKETCHSEGMENT", -40, -20),
    //       new SketchEntityReference("SKETCHSEGMENT", -40, 20));
    public void MirrorSketchEntities(SketchEntityReference mirrorEntity, params SketchEntityReference[] entitiesToMirror)
    {
        EnsureSketchOpen();
        if (!mirrorEntity.IsValid())
        {
            throw new ArgumentException("Mirror operation requires a valid mirrorEntity.", nameof(mirrorEntity));
        }

        if (entitiesToMirror == null || entitiesToMirror.Length == 0)
        {
            throw new ArgumentException("Mirror operation requires at least one entity to mirror.", nameof(entitiesToMirror));
        }

        _activeModel.ClearSelection2(true);

        for (int i = 0; i < entitiesToMirror.Length; i++)
        {
            if (!entitiesToMirror[i].IsValid())
            {
                _activeModel.ClearSelection2(true);
                throw new ArgumentException("One or more mirror target references are invalid.", nameof(entitiesToMirror));
            }

            bool selected = SelectEntityInSketch(entitiesToMirror[i], i > 0, 0);
            if (!selected)
            {
                _activeModel.ClearSelection2(true);
                throw new InvalidOperationException("Failed to select mirror target entity #" + (i + 1) + ".");
            }
        }

        bool mirrorSelected = SelectEntityInSketch(mirrorEntity, true, 0);
        if (!mirrorSelected)
        {
            _activeModel.ClearSelection2(true);
            throw new InvalidOperationException("Failed to select mirror entity.");
        }

        _activeModel.SketchMirror();
        _activeModel.ClearSelection2(true);
    }

    // Create linear sketch step-and-repeat pattern from selected seed entities.
    // Example:
    //   CreateLinearSketchPattern(
    //       3, 1, 20, 0, 0, 90, "",
    //       new SketchEntityReference("SKETCHSEGMENT", -40, -20));
    public void CreateLinearSketchPattern(
        int numX,
        int numY,
        double spacingXmm,
        double spacingYmm,
        double angleXdeg = 0,
        double angleYdeg = 90,
        string deleteInstances = "",
        params SketchEntityReference[] seedEntities)
    {
        EnsureSketchOpen();

        if (numX < 1) throw new ArgumentOutOfRangeException(nameof(numX), "numX must be >= 1.");
        if (numY < 1) throw new ArgumentOutOfRangeException(nameof(numY), "numY must be >= 1.");
        if (spacingXmm < 0) throw new ArgumentOutOfRangeException(nameof(spacingXmm), "spacingXmm must be >= 0.");
        if (spacingYmm < 0) throw new ArgumentOutOfRangeException(nameof(spacingYmm), "spacingYmm must be >= 0.");
        if (seedEntities == null || seedEntities.Length == 0)
        {
            throw new ArgumentException("Linear pattern requires at least one seed entity.", nameof(seedEntities));
        }

        _activeModel.ClearSelection2(true);
        for (int i = 0; i < seedEntities.Length; i++)
        {
            if (!seedEntities[i].IsValid())
            {
                _activeModel.ClearSelection2(true);
                throw new ArgumentException("One or more seed entity references are invalid.", nameof(seedEntities));
            }

            bool selected = SelectEntityInSketch(seedEntities[i], i > 0, 0);
            if (!selected)
            {
                _activeModel.ClearSelection2(true);
                throw new InvalidOperationException("Failed to select linear-pattern seed entity #" + (i + 1) + ".");
            }
        }

        _activeModel.CreateLinearSketchStepAndRepeat(
            numX,
            numY,
            spacingXmm / 1000.0,
            spacingYmm / 1000.0,
            DegreesToRadians(angleXdeg),
            DegreesToRadians(angleYdeg),
            deleteInstances ?? string.Empty);

        _activeModel.ClearSelection2(true);
    }

    // Create circular sketch step-and-repeat pattern from selected seed entities.
    // Example:
    //   CreateCircularSketchPattern(
    //       40, 180, 6, 30, false, "",
    //       new SketchEntityReference("SKETCHSEGMENT", -40, -20));
    public void CreateCircularSketchPattern(
        double arcRadiusMm,
        double arcAngleDeg,
        int patternCount,
        double patternSpacingDeg,
        bool patternRotate = false,
        string deleteInstances = "",
        params SketchEntityReference[] seedEntities)
    {
        EnsureSketchOpen();

        if (arcRadiusMm <= 0) throw new ArgumentOutOfRangeException(nameof(arcRadiusMm), "arcRadiusMm must be > 0.");
        if (patternCount < 1) throw new ArgumentOutOfRangeException(nameof(patternCount), "patternCount must be >= 1.");
        if (seedEntities == null || seedEntities.Length == 0)
        {
            throw new ArgumentException("Circular pattern requires at least one seed entity.", nameof(seedEntities));
        }

        _activeModel.ClearSelection2(true);
        for (int i = 0; i < seedEntities.Length; i++)
        {
            if (!seedEntities[i].IsValid())
            {
                _activeModel.ClearSelection2(true);
                throw new ArgumentException("One or more seed entity references are invalid.", nameof(seedEntities));
            }

            bool selected = SelectEntityInSketch(seedEntities[i], i > 0, 0);
            if (!selected)
            {
                _activeModel.ClearSelection2(true);
                throw new InvalidOperationException("Failed to select circular-pattern seed entity #" + (i + 1) + ".");
            }
        }

        _activeModel.CreateCircularSketchStepAndRepeat(
            arcRadiusMm / 1000.0,
            DegreesToRadians(arcAngleDeg),
            patternCount,
            DegreesToRadians(patternSpacingDeg),
            patternRotate,
            deleteInstances ?? string.Empty);

        _activeModel.ClearSelection2(true);
    }

    // Create sketch fillet using one or two selected entities.
    // firstEntity can be an edge/segment or a sketch point near the target corner.
    public void CreateSketchFillet(
        double radiusMm,
        SketchEntityReference firstEntity,
        SketchEntityReference secondEntity = default,
        bool constrainedCorners = true)
    {
        EnsureSketchOpen();
        if (!firstEntity.IsValid())
        {
            throw new ArgumentException("Sketch fillet requires firstEntity.");
        }

        _activeModel.ClearSelection2(true);
        SelectEntityInSketch(firstEntity, false, 0);
        if (secondEntity.IsValid())
        {
            SelectEntityInSketch(secondEntity, true, 0);
        }

        _activeModel.SketchManager.CreateFillet(radiusMm / 1000.0, constrainedCorners ? 1 : 0);
        _activeModel.ClearSelection2(true);
    }

    // Create sketch chamfer using one or two selected entities.
    // firstEntity can be an edge/segment or a sketch point near the target corner.
    public void CreateSketchChamfer(
        SketchChamferMode chamferMode,
        double distance1Mm,
        double distance2OrAngle,
        SketchEntityReference firstEntity,
        SketchEntityReference secondEntity = default)
    {
        EnsureSketchOpen();
        if (!firstEntity.IsValid())
        {
            throw new ArgumentException("Sketch chamfer requires firstEntity.");
        }

        _activeModel.ClearSelection2(true);
        bool firstSelected = SelectEntityInSketch(firstEntity, false, 0);
        if (!firstSelected)
        {
            throw new InvalidOperationException("Failed to select first chamfer entity in active sketch.");
        }

        if (secondEntity.IsValid())
        {
            bool secondSelected = SelectEntityInSketch(secondEntity, true, 0);
            if (!secondSelected)
            {
                throw new InvalidOperationException("Failed to select second chamfer entity in active sketch.");
            }
        }

        int type = chamferMode == SketchChamferMode.DistanceDistance
            ? (int)swSketchChamferType_e.swSketchChamfer_DistanceDistance
            : (int)swSketchChamferType_e.swSketchChamfer_DistanceAngle;

        double value1 = distance1Mm / 1000.0;
        double value2;
        if (chamferMode == SketchChamferMode.DistanceDistance)
        {
            // If second distance is omitted, use equal-distance chamfer.
            double distance2Mm = double.IsNaN(distance2OrAngle) ? distance1Mm : distance2OrAngle;
            value2 = distance2Mm / 1000.0;
        }
        else
        {
            if (double.IsNaN(distance2OrAngle))
            {
                throw new ArgumentException("DistanceAngle chamfer requires angle value.", nameof(distance2OrAngle));
            }

            value2 = DegreesToRadians(distance2OrAngle);
        }

        _activeModel.SketchManager.CreateChamfer(type, value1, value2);
        _activeModel.ClearSelection2(true);
    }

    // Convenience overload for equal-distance chamfer:
    // when mode is DistanceDistance, distanceMm is applied to both sides.
    public void CreateSketchChamfer(
        SketchChamferMode chamferMode,
        double distanceMm,
        SketchEntityReference firstEntity,
        SketchEntityReference secondEntity = default)
    {
        if (chamferMode != SketchChamferMode.DistanceDistance)
        {
            throw new ArgumentException("This overload is only valid for DistanceDistance chamfer mode.", nameof(chamferMode));
        }

        CreateSketchChamfer(chamferMode, distanceMm, double.NaN, firstEntity, secondEntity);
    }

    // Apply typed sketch relation using one or more selected entities.
    public void ApplySketchRelation(SketchRelationType relationType, params SketchEntityReference[] entities)
    {
        EnsureSketchOpen();
        if (entities == null || entities.Length == 0)
        {
            throw new ArgumentException("At least one entity is required.", nameof(entities));
        }

        _activeModel.ClearSelection2(true);

        for (int i = 0; i < entities.Length; i++)
        {
            if (!entities[i].IsValid())
            {
                throw new ArgumentException("One or more entity references are invalid.", nameof(entities));
            }

            SelectEntityInSketch(entities[i], i > 0, 0);
        }

        _activeModel.SketchAddConstraints(GetConstraintToken(relationType));
        _activeModel.ClearSelection2(true);
    }

    // Select entity in sketch by Name/Type or by coordinate pick (Name can be empty).
    // Coordinates are mm in current sketch space.
    public bool SelectInSketch(string name, string type, double xMm, double yMm, double zMm, bool append, int mark)
    {
        EnsureModelReady();
        return _activeModel.Extension.SelectByID2(name ?? string.Empty, type, xMm / 1000.0, yMm / 1000.0, zMm / 1000.0, append, mark, null, 0);
    }

    // Convenience selector for sketch points by coordinate.
    public bool SelectSketchPoint(double xMm, double yMm, bool append = false, int mark = 0)
    {
        return SelectInSketch(string.Empty, "SKETCHPOINT", xMm, yMm, 0, append, mark);
    }

    // Add one SOLIDWORKS sketch constraint token (example: sgCOINCIDENT, sgHORIZONTAL).
    public void AddSketchConstraint(string constraintToken)
    {
        EnsureSketchOpen();
        _activeModel.SketchAddConstraints(constraintToken);
    }

    // Add SmartDimension at provided text location (mm).
    public void AddDimension(double textXmm, double textYmm, double textZmm = 0, double? val = null)
    {
        EnsureSketchOpen();
        DisplayDimension displayDim = (DisplayDimension)_activeModel.AddDimension2(textXmm / 1000.0, textYmm / 1000.0, textZmm / 1000.0);
        if (displayDim != null && val.HasValue)
        {
            Dimension swDim = displayDim.GetDimension();
            swDim.SystemValue = val.Value / 1000.0;
        }
    }

    // Add SmartDimension by selecting two sketch references from 2D coordinates,
    // then placing the dimension text at the provided offset position.
    public void AddDimension(
        double firstXmm,
        double firstYmm,
        double secondXmm,
        double secondYmm,
        double textXmm,
        double textYmm,
        double textZmm = 0,
        string firstType = "",
        string secondType = "",
        double? val = null)
    {
        EnsureSketchOpen();

        _activeModel.ClearSelection2(true);

        bool firstSelected = SelectDimensionReference(firstXmm, firstYmm, false, 0, firstType, preferPoint: true);
        if (!firstSelected)
        {
            throw new InvalidOperationException(
                "Failed to select first reference for dimension at (" + firstXmm + ", " + firstYmm + ") mm.");
        }

        bool secondSelected = SelectDimensionReference(secondXmm, secondYmm, true, 0, secondType, preferPoint: true);
        if (!secondSelected)
        {
            _activeModel.ClearSelection2(true);
            throw new InvalidOperationException(
                "Failed to select second reference for dimension at (" + secondXmm + ", " + secondYmm + ") mm.");
        }

        DisplayDimension displayDim = (DisplayDimension)_activeModel.AddDimension2(textXmm / 1000.0, textYmm / 1000.0, textZmm / 1000.0);
        _activeModel.ClearSelection2(true);
        if (displayDim == null)
        {
            throw new InvalidOperationException("Smart Dimension creation failed after selecting both references.");
        }

        if (val.HasValue)
        {
            Dimension swDim = displayDim.GetDimension();
            swDim.SystemValue = val.Value / 1000.0;
        }
    }

    // Add SmartDimension from a single selected entity/point, then place text.
    // Useful for line length, circle diameter/radius, arc dimensions, etc.
    public void AddDimension(
        double entityXmm,
        double entityYmm,
        double textXmm,
        double textYmm,
        double textZmm = 0,
        string entityType = "",
        double? val = null)
    {
        EnsureSketchOpen();

        _activeModel.ClearSelection2(true);

        // For single-entity dimensions, prefer segment/arc picks over point picks.
        bool selected = SelectDimensionReference(entityXmm, entityYmm, false, 0, entityType, preferPoint: false);
        if (!selected)
        {
            throw new InvalidOperationException(
                "Failed to select dimension entity at (" + entityXmm + ", " + entityYmm + ") mm.");
        }

        DisplayDimension displayDim = (DisplayDimension)_activeModel.AddDimension2(textXmm / 1000.0, textYmm / 1000.0, textZmm / 1000.0);
        _activeModel.ClearSelection2(true);
        if (displayDim == null)
        {
            throw new InvalidOperationException("Smart Dimension creation failed for single-entity selection.");
        }

        if (val.HasValue)
        {
            Dimension swDim = displayDim.GetDimension();
            swDim.SystemValue = val.Value / 1000.0;
        }
    }

    // Exit sketch mode.
    public void EndSketch()
    {
        EnsureSketchOpen();
        _activeModel.SketchManager.InsertSketch(true);
        _isSketchOpen = false;
    }

    // Create extrusion feature from current sketch result.
    // depthMm is in mm, midPlane controls end condition, isCut toggles cut-extrude.
    public void Extrude(double depthMm, bool midPlane = false, bool isCut = false)
    {
        EnsureModelReady();
        if (_isSketchOpen) EndSketch();

        int endCondition = midPlane
            ? (int)swEndConditions_e.swEndCondMidPlane
            : (int)swEndConditions_e.swEndCondBlind;

        Feature createdExtrude = _activeModel.FeatureManager.FeatureExtrusion2(
            !isCut, false, false,
            endCondition, 0,
            depthMm / 1000.0, 0,
            false, false, false, false,
            0, 0, false, false, false, false,
            true, true, true, 0, 0, false);

        // Create helper reference planes after extrusion:
        // - Non-cylindrical result: create all face-direction planes.
        // - Cylindrical result: create Top/Bottom only.
        if (createdExtrude != null)
        {
            TryCreateFaceDirectionPlanesAfterExtrude();
        }
    }

    // Rebuild active part.
    public void Rebuild()
    {
        EnsureModelReady();
        _activeModel.ForceRebuild3(false);
    }

    // Save active part to configured output folder.
    public string SavePart()
    {
        EnsureModelReady();
        if (string.IsNullOrWhiteSpace(_activePartName) || string.IsNullOrWhiteSpace(_outputFolder))
        {
            throw new InvalidOperationException("No active part context to save.");
        }

        string fullPath = Path.Combine(_outputFolder, _activePartName + ".SLDPRT");
        _activeModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        return fullPath;
    }

    // Save and optionally close active document.
    public string SaveAndClose(bool closeDocument)
    {
        string savedPath = SavePart();
        if (closeDocument)
        {
            string title = _activeModel.GetTitle();
            _swApp.CloseDoc(title);
            _activeModel = null;
            _isSketchOpen = false;
        }

        return savedPath;
    }

    // Return currently active model (advanced use).
    public ModelDoc2 GetActiveModel()
    {
        EnsureModelReady();
        return _activeModel;
    }

    // Ensure active model exists.
    private void EnsureModelReady()
    {
        if (_activeModel == null) throw new InvalidOperationException("No active sketch part. Call BeginPartSketch first.");
    }

    // Ensure active sketch mode exists.
    private void EnsureSketchOpen()
    {
        EnsureModelReady();
        if (!_isSketchOpen) throw new InvalidOperationException("Sketch is not open. Call BeginPartSketch or reopen sketch first.");
    }

    // Select entity in sketch using a typed entity reference.
    private bool SelectEntityInSketch(SketchEntityReference reference, bool append, int mark)
    {
        return SelectInSketch(reference.Name, reference.Type, reference.XMm, reference.YMm, reference.ZMm, append, mark);
    }

    // Select a reference usable by Smart Dimension using either an explicit type
    // or a best-effort fallback order for common sketch entities.
    private bool SelectDimensionReference(double xMm, double yMm, bool append, int mark, string explicitType = "", bool preferPoint = true)
    {
        if (!string.IsNullOrWhiteSpace(explicitType))
        {
            return SelectInSketch(string.Empty, explicitType, xMm, yMm, 0, append, mark);
        }

        string[] fallbackTypes = preferPoint
            ? new[] { "SKETCHPOINT", "SKETCHSEGMENT", "SKETCHARC" }
            : new[] { "SKETCHSEGMENT", "SKETCHARC", "SKETCHPOINT" };

        for (int i = 0; i < fallbackTypes.Length; i++)
        {
            if (SelectInSketch(string.Empty, fallbackTypes[i], xMm, yMm, 0, append, mark))
            {
                return true;
            }
        }

        return false;
    }

    // Map relation enum to SOLIDWORKS constraint token.
    private static string GetConstraintToken(SketchRelationType relationType)
    {
        switch (relationType)
        {
            case SketchRelationType.Horizontal:
                return "sgHORIZONTAL";
            case SketchRelationType.Vertical:
                return "sgVERTICAL";
            case SketchRelationType.Parallel:
                return "sgPARALLEL";
            case SketchRelationType.Perpendicular:
                return "sgPERPENDICULAR";
            case SketchRelationType.Tangent:
                return "sgTANGENT";
            case SketchRelationType.Equal:
                return "sgEQUAL";
            case SketchRelationType.Symmetric:
                return "sgSYMMETRIC";
            case SketchRelationType.Coincident:
                return "sgCOINCIDENT";
            case SketchRelationType.Collinear:
                return "sgCOLINEAR";
            default:
                throw new ArgumentOutOfRangeException(nameof(relationType), relationType, "Unsupported relation type.");
        }
    }

    // Convert angle in degrees to radians for SOLIDWORKS APIs.
    private static double DegreesToRadians(double degrees)
    {
        return degrees * Math.PI / 180.0;
    }

    // Best-effort segment count of currently active sketch.
    // Returns -1 when the count cannot be determined.
    private int GetActiveSketchSegmentCount()
    {
        Sketch activeSketch = _activeModel.SketchManager.ActiveSketch;
        if (activeSketch == null) return -1;

        object segmentsObj = activeSketch.GetSketchSegments();
        object[] segments = segmentsObj as object[];
        return segments?.Length ?? 0;
    }

    // Build rectangle manually as 4 lines.
    // Corner mode: (x1,y1) and (x2,y2) are opposite corners.
    // Center mode: (x1,y1) is center, (x2,y2) is one corner.
    private void CreateRectangleByLines(SketchRectangleType rectangleType, double x1Mm, double y1Mm, double x2Mm, double y2Mm)
    {
        double minX;
        double maxX;
        double minY;
        double maxY;
        GetRectangleBounds(rectangleType, x1Mm, y1Mm, x2Mm, y2Mm, out minX, out maxX, out minY, out maxY);

        bool previousAddToDb = _activeModel.SketchManager.AddToDB;
        bool previousDisplay = _activeModel.SketchManager.DisplayWhenAdded;

        try
        {
            // Disable inference/snap while creating fallback geometry to avoid skewed/merged segments.
            _activeModel.SketchManager.AddToDB = true;
            _activeModel.SketchManager.DisplayWhenAdded = false;

            object l1 = _activeModel.SketchManager.CreateLine(minX / 1000.0, minY / 1000.0, 0, maxX / 1000.0, minY / 1000.0, 0);
            object l2 = _activeModel.SketchManager.CreateLine(maxX / 1000.0, minY / 1000.0, 0, maxX / 1000.0, maxY / 1000.0, 0);
            object l3 = _activeModel.SketchManager.CreateLine(maxX / 1000.0, maxY / 1000.0, 0, minX / 1000.0, maxY / 1000.0, 0);
            object l4 = _activeModel.SketchManager.CreateLine(minX / 1000.0, maxY / 1000.0, 0, minX / 1000.0, minY / 1000.0, 0);

            if (l1 == null || l2 == null || l3 == null || l4 == null)
            {
                throw new InvalidOperationException("Fallback rectangle line creation returned null segment(s).");
            }
        }
        finally
        {
            _activeModel.SketchManager.AddToDB = previousAddToDb;
            _activeModel.SketchManager.DisplayWhenAdded = previousDisplay;
            _activeModel.GraphicsRedraw2();
        }
    }

    // Compute rectangle min/max bounds in mm for both Corner and Center input modes.
    private static void GetRectangleBounds(
        SketchRectangleType rectangleType,
        double x1Mm,
        double y1Mm,
        double x2Mm,
        double y2Mm,
        out double minX,
        out double maxX,
        out double minY,
        out double maxY)
    {
        if (rectangleType == SketchRectangleType.Center)
        {
            double halfWidth = Math.Abs(x2Mm - x1Mm);
            double halfHeight = Math.Abs(y2Mm - y1Mm);
            minX = x1Mm - halfWidth;
            maxX = x1Mm + halfWidth;
            minY = y1Mm - halfHeight;
            maxY = y1Mm + halfHeight;
            return;
        }

        minX = Math.Min(x1Mm, x2Mm);
        maxX = Math.Max(x1Mm, x2Mm);
        minY = Math.Min(y1Mm, y2Mm);
        maxY = Math.Max(y1Mm, y2Mm);
    }

    // Create helper planes after extrusion based on resulting solid shape.
    private void TryCreateFaceDirectionPlanesAfterExtrude()
    {
        double minX, minY, minZ, maxX, maxY, maxZ;
        if (!TryGetFirstSolidBodyBox(out minX, out minY, out minZ, out maxX, out maxY, out maxZ)) return;

        bool hasCylinder = HasCylindricalSurface();

        // For circular/cylindrical results, create only Top/Bottom helper planes.
        if (hasCylinder)
        {
            CreateOffsetPlaneFromBaseCandidates(
                new[] { "Top Plane", "Top", "Ebene Oben", "Oben" },
                Math.Abs(maxZ),
                false,
                "Face_Top");
            CreateOffsetPlaneFromBaseCandidates(
                new[] { "Top Plane", "Top", "Ebene Oben", "Oben" },
                Math.Abs(minZ),
                true,
                "Face_Bottom");
            return;
        }

        // For rectangular/non-cylindrical results, create full directional planes.
        CreateOffsetPlaneFromBaseCandidates(
            new[] { "Front Plane", "Front", "Ebene Vorne", "Vorne" },
            Math.Abs(maxY),
            false,
            "Face_Front");
        CreateOffsetPlaneFromBaseCandidates(
            new[] { "Front Plane", "Front", "Ebene Vorne", "Vorne" },
            Math.Abs(minY),
            true,
            "Face_Back");
        CreateOffsetPlaneFromBaseCandidates(
            new[] { "Right Plane", "Right", "Ebene Rechts", "Rechts" },
            Math.Abs(maxX),
            false,
            "Face_Right");
        CreateOffsetPlaneFromBaseCandidates(
            new[] { "Right Plane", "Right", "Ebene Rechts", "Rechts" },
            Math.Abs(minX),
            true,
            "Face_Left");
        CreateOffsetPlaneFromBaseCandidates(
            new[] { "Top Plane", "Top", "Ebene Oben", "Oben" },
            Math.Abs(maxZ),
            false,
            "Face_Top");
        CreateOffsetPlaneFromBaseCandidates(
            new[] { "Top Plane", "Top", "Ebene Oben", "Oben" },
            Math.Abs(minZ),
            true,
            "Face_Bottom");
    }

    // Detect whether the first visible solid body contains at least one cylindrical face.
    private bool HasCylindricalSurface()
    {
        PartDoc partDoc = _activeModel as PartDoc;
        if (partDoc == null) return false;

        object bodyArrayObj = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
        if (bodyArrayObj == null) return false;

        object[] bodies = bodyArrayObj as object[];
        if (bodies == null || bodies.Length == 0) return false;

        for (int i = 0; i < bodies.Length; i++)
        {
            Body2 body = bodies[i] as Body2;
            if (body == null) continue;

            object faceArrayObj = body.GetFaces();
            object[] faces = faceArrayObj as object[];
            if (faces == null) continue;

            for (int j = 0; j < faces.Length; j++)
            {
                Face2 face = faces[j] as Face2;
                if (face == null) continue;

                Surface surface = face.GetSurface() as Surface;
                if (surface == null || !surface.IsCylinder()) continue;
                return true;
            }
        }

        return false;
    }

    // Get the bounding box of the first visible solid body in part coordinates (meters).
    private bool TryGetFirstSolidBodyBox(out double minX, out double minY, out double minZ, out double maxX, out double maxY, out double maxZ)
    {
        minX = minY = minZ = maxX = maxY = maxZ = 0;

        PartDoc partDoc = _activeModel as PartDoc;
        if (partDoc == null) return false;

        object bodyArrayObj = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
        if (bodyArrayObj == null) return false;

        object[] bodies = bodyArrayObj as object[];
        if (bodies == null || bodies.Length == 0) return false;

        Body2 body = bodies[0] as Body2;
        if (body == null) return false;

        object boxObj = body.GetBodyBox();
        double[] box = boxObj as double[];
        if (box == null || box.Length < 6) return false;

        minX = box[0];
        minY = box[1];
        minZ = box[2];
        maxX = box[3];
        maxY = box[4];
        maxZ = box[5];
        return true;
    }

    // Create one offset plane from a localized base-plane name set.
    private void CreateOffsetPlaneFromBaseCandidates(string[] basePlaneCandidates, double distanceMeters, bool reverseDirection, string newPlaneName)
    {
        // Reuse existing helper plane when available.
        if (SelectReferencePlaneByNameOrPrefix(newPlaneName, false))
        {
            _activeModel.ClearSelection2(true);
            return;
        }

        _activeModel.ClearSelection2(true);
        bool selected = SelectPartPlane(_activeModel, basePlaneCandidates, false);
        if (!selected) return;

        Feature plane = _activeModel.CreatePlaneAtOffset3(distanceMeters, reverseDirection, false);
        if (plane == null) return;

        // Naming is best-effort; ignore collisions/localized behavior.
        try
        {
            plane.Name = newPlaneName;
        }
        catch
        {
            try
            {
                plane.Name = newPlaneName + "_Auto";
            }
            catch
            {
            }
        }
    }

    // Select a reference plane by exact name or prefix name (handles auto suffixes like <1>).
    private bool SelectReferencePlaneByNameOrPrefix(string requestedName, bool append)
    {
        if (string.IsNullOrWhiteSpace(requestedName)) return false;

        Feature feat = _activeModel.FirstFeature();
        while (feat != null)
        {
            if (string.Equals(feat.GetTypeName2(), "RefPlane", StringComparison.OrdinalIgnoreCase))
            {
                string featName = feat.Name ?? string.Empty;
                bool nameMatch =
                    string.Equals(featName, requestedName, StringComparison.OrdinalIgnoreCase) ||
                    featName.StartsWith(requestedName + "<", StringComparison.OrdinalIgnoreCase) ||
                    featName.StartsWith(requestedName + "_", StringComparison.OrdinalIgnoreCase);

                if (nameMatch)
                {
                    return feat.Select2(append, 0);
                }
            }

            feat = feat.GetNextFeature();
        }

        return false;
    }

    // Map helper face-plane aliases to default datum planes.
    // This prevents hard failures when helper planes are not present yet.
    private static string[] GetFacePlaneAliasCandidates(string planeName)
    {
        if (string.IsNullOrWhiteSpace(planeName)) return Array.Empty<string>();

        switch (planeName.Trim().ToLowerInvariant())
        {
            case "face_top":
            case "face_bottom":
                return new[] { "Top Plane", "Top", "Ebene Oben", "Oben" };
            case "face_front":
            case "face_back":
                return new[] { "Front Plane", "Front", "Ebene Vorne", "Vorne" };
            case "face_right":
            case "face_left":
                return new[] { "Right Plane", "Right", "Ebene Rechts", "Rechts" };
            default:
                return Array.Empty<string>();
        }
    }

    // Resolve localized plane names from requested option.
    private static string[] GetSketchPlaneCandidates(SketchPlaneName sketchPlane)
    {
        switch (sketchPlane)
        {
            case SketchPlaneName.Top:
                return new[] { "Top Plane", "Top", "Ebene Oben", "Oben" };
            case SketchPlaneName.Right:
            case SketchPlaneName.Side:
                return new[] { "Right Plane", "Right", "Ebene Rechts", "Rechts" };
            case SketchPlaneName.Front:
            default:
                return new[] { "Front Plane", "Front", "Ebene Vorne", "Vorne" };
        }
    }

    // Select plane by known candidate names.
    private static bool SelectPartPlane(ModelDoc2 partModel, string[] planeNames, bool append)
    {
        for (int i = 0; i < planeNames.Length; i++)
        {
            bool selected = partModel.Extension.SelectByID2(planeNames[i], "PLANE", 0, 0, 0, append, 0, null, 0);
            if (selected) return true;
        }

        return false;
    }

    // Fallback: pick first reference plane in feature tree.
    private static bool SelectFirstReferencePlane(ModelDoc2 partModel, bool append)
    {
        Feature feat = partModel.FirstFeature();
        while (feat != null)
        {
            if (string.Equals(feat.GetTypeName2(), "RefPlane", StringComparison.OrdinalIgnoreCase))
            {
                if (feat.Select2(append, 0)) return true;
            }

            feat = feat.GetNextFeature();
        }

        return false;
    }
}
