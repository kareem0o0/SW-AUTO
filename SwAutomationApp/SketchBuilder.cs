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

// Point creation variants supported by SketchBuilder.
public enum SketchPointType
{
    Standard,
    Intersection,
    Midpoint
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
    public string Name { get; }

    public string Type { get; }

    public double XMm { get; }

    public double YMm { get; }

    public double ZMm { get; }

    public SketchEntityReference(string type, double xMm, double yMm, double zMm = 0, string name = "")
    {
        Type = type ?? string.Empty;
        XMm = xMm;
        YMm = yMm;
        ZMm = zMm;
        Name = name ?? string.Empty;
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

        // Build rectangle by requested variant.
        switch (rectangleType)
        {
            case SketchRectangleType.Center:
                _activeModel.SketchManager.CreateCenterRectangle(x1, y1, 0, x2, y2, 0);
                break;
            case SketchRectangleType.Corner:
            default:
                _activeModel.SketchManager.CreateCornerRectangle(x1, y1, 0, x2, y2, 0);
                break;
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

    // Create point entity with requested point workflow.
    public void CreatePoint(
        SketchPointType pointType,
        double xMm,
        double yMm,
        SketchEntityReference firstReference = default,
        SketchEntityReference secondReference = default)
    {
        EnsureSketchOpen();

        _activeModel.SketchManager.CreatePoint(xMm / 1000.0, yMm / 1000.0, 0);

        switch (pointType)
        {
            case SketchPointType.Intersection:
                if (!firstReference.IsValid() || !secondReference.IsValid())
                {
                    throw new ArgumentException("Intersection point requires firstReference and secondReference.");
                }

                ApplyPointCoincidentToEntity(xMm, yMm, firstReference);
                ApplyPointCoincidentToEntity(xMm, yMm, secondReference);
                break;

            case SketchPointType.Midpoint:
                if (!firstReference.IsValid())
                {
                    throw new ArgumentException("Midpoint requires firstReference for the target line.");
                }

                _activeModel.ClearSelection2(true);
                SelectSketchPoint(xMm, yMm, false, 0);
                SelectEntityInSketch(firstReference, true, 0);
                _activeModel.SketchAddConstraints("sgMIDPOINT");
                _activeModel.ClearSelection2(true);
                break;

            case SketchPointType.Standard:
            default:
                break;
        }
    }

    // Create standard sketch point in mm coordinates.
    public void CreatePoint(double xMm, double yMm)
    {
        CreatePoint(SketchPointType.Standard, xMm, yMm);
    }

    // Create sketch fillet between two selected entities.
    public void CreateSketchFillet(double radiusMm, SketchEntityReference firstEntity, SketchEntityReference secondEntity, bool constrainedCorners = true)
    {
        EnsureSketchOpen();
        if (!firstEntity.IsValid() || !secondEntity.IsValid())
        {
            throw new ArgumentException("Sketch fillet requires firstEntity and secondEntity.");
        }

        _activeModel.ClearSelection2(true);
        SelectEntityInSketch(firstEntity, false, 0);
        SelectEntityInSketch(secondEntity, true, 0);
        _activeModel.SketchManager.CreateFillet(radiusMm / 1000.0, constrainedCorners ? 1 : 0);
        _activeModel.ClearSelection2(true);
    }

    // Create sketch chamfer between two selected entities.
    public void CreateSketchChamfer(
        SketchChamferMode chamferMode,
        double distance1Mm,
        double distance2OrAngle,
        SketchEntityReference firstEntity,
        SketchEntityReference secondEntity)
    {
        EnsureSketchOpen();
        if (!firstEntity.IsValid() || !secondEntity.IsValid())
        {
            throw new ArgumentException("Sketch chamfer requires firstEntity and secondEntity.");
        }

        _activeModel.ClearSelection2(true);
        SelectEntityInSketch(firstEntity, false, 0);
        SelectEntityInSketch(secondEntity, true, 0);

        int type = chamferMode == SketchChamferMode.DistanceDistance
            ? (int)swSketchChamferType_e.swSketchChamfer_DistanceDistance
            : (int)swSketchChamferType_e.swSketchChamfer_DistanceAngle;

        double value1 = distance1Mm / 1000.0;
        double value2 = chamferMode == SketchChamferMode.DistanceDistance
            ? distance2OrAngle / 1000.0
            : DegreesToRadians(distance2OrAngle);

        _activeModel.SketchManager.CreateChamfer(type, value1, value2);
        _activeModel.ClearSelection2(true);
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
    public void AddDimension(double textXmm, double textYmm, double textZmm = 0)
    {
        EnsureSketchOpen();
        _activeModel.AddDimension2(textXmm / 1000.0, textYmm / 1000.0, textZmm / 1000.0);
    }

    // Exit sketch mode.
    public void EndSketch()
    {
        EnsureSketchOpen();
        _activeModel.SketchManager.InsertSketch(true);
        _isSketchOpen = false;
    }

    // Create extrusion feature from current sketch result.
    // depthMm is in mm, midPlane controls end condition.
    public void Extrude(double depthMm, bool midPlane = false)
    {
        EnsureModelReady();
        if (_isSketchOpen) EndSketch();

        int endCondition = midPlane
            ? (int)swEndConditions_e.swEndCondMidPlane
            : (int)swEndConditions_e.swEndCondBlind;

        _activeModel.FeatureManager.FeatureExtrusion2(
            true, false, false,
            endCondition, 0,
            depthMm / 1000.0, 0,
            false, false, false, false,
            0, 0, false, false, false, false,
            true, true, true, 0, 0, false);
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

    // Apply coincident relation between sketch point and another sketch entity.
    private void ApplyPointCoincidentToEntity(double xMm, double yMm, SketchEntityReference entityReference)
    {
        _activeModel.ClearSelection2(true);
        SelectSketchPoint(xMm, yMm, false, 0);
        SelectEntityInSketch(entityReference, true, 0);
        _activeModel.SketchAddConstraints("sgCOINCIDENT");
        _activeModel.ClearSelection2(true);
    }

    // Select entity in sketch using a typed entity reference.
    private bool SelectEntityInSketch(SketchEntityReference reference, bool append, int mark)
    {
        return SelectInSketch(reference.Name, reference.Type, reference.XMm, reference.YMm, reference.ZMm, append, mark);
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
