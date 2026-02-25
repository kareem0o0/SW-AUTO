using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation;

public enum SketchRectangleType { Center, Corner }
public enum SketchCircleType { CenterRadius, CenterPoint }
public enum SketchLineType { Standard, Construction, Centerline }
public enum SketchRelationType { Horizontal, Vertical, Parallel, Perpendicular, Tangent, Equal, Symmetric, Coincident, Collinear }
public enum SketchPlaneName0 { Front, Top, Right, Side }
public enum SketchChamferMode { DistanceAngle, DistanceDistance }

// KEEP THIS - Your code depends on it
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
        Type = string.IsNullOrWhiteSpace(type) ? DefaultType : type;
        XMm = xMm;
        YMm = yMm;
        ZMm = zMm;
        Name = name ?? string.Empty;
    }

    public SketchEntityReference(double xMm, double yMm, double zMm = 0, string name = "")
        : this(DefaultType, xMm, yMm, zMm, name)
    {
    }

    public bool IsValid() => !string.IsNullOrWhiteSpace(Type);
}

public sealed class SketchBuilder
{
    private readonly SldWorks _swApp;
    private ModelDoc2 _activeModel;
    private SketchManager _sketchManager;
    private string _activePartName = string.Empty;
    private string _outputFolder = string.Empty;
    private bool _isSketchOpen;

    private const double MmToMeters = 0.001;
    private const string SketchPointType = "SKETCHPOINT";
    private const string SketchSegmentType = "SKETCHSEGMENT";
    private const string PlaneType = "PLANE";

    public SketchBuilder(SldWorks swApp)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
    }

    public void BeginPartSketch(string partName, string outputFolder, SketchPlaneName0 sketchPlane = SketchPlaneName0.Front)
    {
        if (string.IsNullOrWhiteSpace(partName))
            throw new ArgumentException("Part name is required.", nameof(partName));
        if (string.IsNullOrWhiteSpace(outputFolder))
            throw new ArgumentException("Output folder is required.", nameof(outputFolder));

        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        _activeModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
        _sketchManager = _activeModel.SketchManager;
        _activePartName = partName;
        _outputFolder = outputFolder;

        StartSketchOnPlane(GetPlaneName(sketchPlane));
    }

    public void BeginSketch(SketchPlaneName0 sketchPlane = SketchPlaneName0.Front, ModelDoc2 partModel = null)
    {
        if (partModel != null)
        {
            _activeModel = partModel;
            _sketchManager = _activeModel.SketchManager;
        }

        EnsureModelReady();
        CloseOpenSketch();
        StartSketchOnPlane(GetPlaneName(sketchPlane));
    }

    // KEEP THIS - Your code uses it with SketchEntityReference
    public void BeginSketch(SketchEntityReference sketchTargetReference, ModelDoc2 partModel = null)
    {
        if (partModel != null)
        {
            _activeModel = partModel;
            _sketchManager = _activeModel.SketchManager;
        }

        EnsureModelReady();
        
        if (!sketchTargetReference.IsValid())
            throw new ArgumentException("Invalid sketch target reference", nameof(sketchTargetReference));

        CloseOpenSketch();
        _activeModel.ClearSelection2(true);

        bool selected = _activeModel.Extension.SelectByID2(
            sketchTargetReference.Name ?? string.Empty,
            sketchTargetReference.Type,
            sketchTargetReference.XMm * MmToMeters,
            sketchTargetReference.YMm * MmToMeters,
            sketchTargetReference.ZMm * MmToMeters,
            false, 0, null, 0);

        if (!selected)
            throw new InvalidOperationException($"Failed to select sketch target: {sketchTargetReference.Type}");

        if (string.Equals(sketchTargetReference.Type, "FACE", StringComparison.OrdinalIgnoreCase))
        {
            var selMgr = (SelectionMgr)_activeModel.SelectionManager;
            var selectedFace = selMgr?.GetSelectedObject6(1, -1) as Face2;
            var surface = selectedFace?.GetSurface() as Surface;
            
            if (surface == null || !surface.IsPlane())
            {
                _activeModel.ClearSelection2(true);
                throw new InvalidOperationException("Selected face is not planar.");
            }
        }

        _sketchManager.InsertSketch(true);
        _isSketchOpen = true;
    }

    public void CreateRectangle(SketchRectangleType rectangleType, double x1Mm, double y1Mm, double x2Mm, double y2Mm)
    {
        EnsureSketchOpen();

        double x1 = x1Mm * MmToMeters;
        double y1 = y1Mm * MmToMeters;
        double x2 = x2Mm * MmToMeters;
        double y2 = y2Mm * MmToMeters;

        object result = rectangleType == SketchRectangleType.Center
            ? _sketchManager.CreateCenterRectangle(x1, y1, 0, x2, y2, 0)
            : _sketchManager.CreateCornerRectangle(x1, y1, 0, x2, y2, 0);

        if (result == null)
            throw new InvalidOperationException($"Failed to create {rectangleType} rectangle");
    }

    public void CreateCircle(SketchCircleType circleType, double centerXmm, double centerYmm, double value1Mm, double value2Mm = 0)
    {
        EnsureSketchOpen();

        double cx = centerXmm * MmToMeters;
        double cy = centerYmm * MmToMeters;
        double v1 = value1Mm * MmToMeters;
        double v2 = value2Mm * MmToMeters;

        if (circleType == SketchCircleType.CenterRadius)
            _sketchManager.CreateCircleByRadius(cx, cy, 0, v1);
        else
            _sketchManager.CreateCircle(cx, cy, 0, v1, v2, 0);
    }

    public void CreateLine(SketchLineType lineType, double startXmm, double startYmm, double endXmm, double endYmm)
    {
        EnsureSketchOpen();

        double startX = startXmm * MmToMeters;
        double startY = startYmm * MmToMeters;
        double endX = endXmm * MmToMeters;
        double endY = endYmm * MmToMeters;

        switch (lineType)
        {
            case SketchLineType.Centerline:
                _sketchManager.CreateCenterLine(startX, startY, 0, endX, endY, 0);
                break;
            case SketchLineType.Construction:
                _sketchManager.CreateLine(startX, startY, 0, endX, endY, 0);
                _activeModel.ClearSelection2(true);
                
                double midX = (startX + endX) / 2.0;
                double midY = (startY + endY) / 2.0;
                
                if (_activeModel.Extension.SelectByID2(string.Empty, SketchSegmentType, midX, midY, 0, false, 0, null, 0))
                    _sketchManager.CreateConstructionGeometry();
                
                _activeModel.ClearSelection2(true);
                break;
            default:
                _sketchManager.CreateLine(startX, startY, 0, endX, endY, 0);
                break;
        }
    }

    public void CreateLine(double startXmm, double startYmm, double endXmm, double endYmm)
    {
        CreateLine(SketchLineType.Standard, startXmm, startYmm, endXmm, endYmm);
    }

    public void CreatePoint(double xMm, double yMm)
    {
        EnsureSketchOpen();
        _sketchManager.CreatePoint(xMm * MmToMeters, yMm * MmToMeters, 0);
    }

    // KEEP THIS - Your code uses it
    public void CreateSketchTrim(double xMm, double yMm, double zMm = double.NaN, int trimMode = 0)
    {
        EnsureSketchOpen();
        
        double z = double.IsNaN(zMm) ? 0 : zMm * MmToMeters;
        
        if (trimMode == 0 || trimMode == 3)
        {
            _activeModel.ClearSelection2(true);
            _activeModel.Extension.SelectByID2(string.Empty, SketchSegmentType, xMm * MmToMeters, yMm * MmToMeters, z, false, 0, null, 0);
        }

        _sketchManager.SketchTrim(trimMode, xMm * MmToMeters, yMm * MmToMeters, z);
        _activeModel.ClearSelection2(true);
    }

    // KEEP THIS - Your code uses it
    public void ApplySketchRelation(SketchRelationType relationType, params SketchEntityReference[] entities)
    {
        EnsureSketchOpen();
        
        if (entities == null || entities.Length == 0)
            throw new ArgumentException("At least one entity required", nameof(entities));

        _activeModel.ClearSelection2(true);

        for (int i = 0; i < entities.Length; i++)
        {
            if (!entities[i].IsValid())
                throw new ArgumentException($"Entity {i + 1} is invalid");

            bool selected = _activeModel.Extension.SelectByID2(
                entities[i].Name ?? string.Empty,
                entities[i].Type,
                entities[i].XMm * MmToMeters,
                entities[i].YMm * MmToMeters,
                entities[i].ZMm * MmToMeters,
                i > 0, 0, null, 0);

            if (!selected)
            {
                _activeModel.ClearSelection2(true);
                throw new InvalidOperationException($"Failed to select entity {i + 1}");
            }
        }

        _activeModel.SketchAddConstraints(GetConstraintToken(relationType));
        _activeModel.ClearSelection2(true);
    }

    // KEEP THIS - Convenience overload for point-based relations
    public void ApplySketchRelation(SketchRelationType relationType, double firstXmm, double firstYmm, double secondXmm, double secondYmm)
    {
        EnsureSketchOpen();

        _activeModel.ClearSelection2(true);

        bool firstSelected = _activeModel.Extension.SelectByID2(
            string.Empty, SketchPointType, 
            firstXmm * MmToMeters, firstYmm * MmToMeters, 0, 
            false, 0, null, 0);
            
        if (!firstSelected)
            throw new InvalidOperationException($"Failed to select first point at ({firstXmm}, {firstYmm})");

        bool secondSelected = _activeModel.Extension.SelectByID2(
            string.Empty, SketchPointType, 
            secondXmm * MmToMeters, secondYmm * MmToMeters, 0, 
            true, 0, null, 0);
            
        if (!secondSelected)
        {
            _activeModel.ClearSelection2(true);
            throw new InvalidOperationException($"Failed to select second point at ({secondXmm}, {secondYmm})");
        }

        _activeModel.SketchAddConstraints(GetConstraintToken(relationType));
        _activeModel.ClearSelection2(true);
    }

    // KEEP THIS - Your code uses it
    public bool DeleteSketchEntityAt(double xMm, double yMm, double zMm = double.NaN, string explicitType = "")
    {
        EnsureSketchOpen();
        
        double z = double.IsNaN(zMm) ? 0 : zMm * MmToMeters;
        _activeModel.ClearSelection2(true);

        bool selected = false;
        
        if (!string.IsNullOrWhiteSpace(explicitType))
        {
            selected = _activeModel.Extension.SelectByID2(
                string.Empty, explicitType, 
                xMm * MmToMeters, yMm * MmToMeters, z, 
                false, 0, null, 0);
        }
        else
        {
            string[] types = { SketchSegmentType, "SKETCHARC", SketchPointType };
            foreach (string type in types)
            {
                selected = _activeModel.Extension.SelectByID2(
                    string.Empty, type, 
                    xMm * MmToMeters, yMm * MmToMeters, z, 
                    false, 0, null, 0);
                if (selected) break;
            }
        }

        if (!selected) return false;

        try
        {
            _activeModel.EditDelete();
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            _activeModel.ClearSelection2(true);
        }
    }

    public void AddDimension(double textXmm, double textYmm, double? valueMm = null)
    {
        EnsureSketchOpen();
        
        var displayDim = (DisplayDimension)_activeModel.AddDimension2(
            textXmm * MmToMeters, textYmm * MmToMeters, 0);
        
        if (displayDim != null && valueMm.HasValue)
        {
            var swDim = displayDim.GetDimension();
            swDim.SystemValue = valueMm.Value * MmToMeters;
        }
    }

    public void AddDimension(double firstXmm, double firstYmm, double secondXmm, double secondYmm, 
                        double textXmm, double textYmm, double? valueMm = null)
{
    EnsureSketchOpen();
    _activeModel.ClearSelection2(true);

    // Select first point
    bool firstSelected = _activeModel.Extension.SelectByID2(
        string.Empty, "SKETCHPOINT", 
        firstXmm * MmToMeters, firstYmm * MmToMeters, 0, 
        false, 0, null, 0);
        
    if (!firstSelected)
        throw new InvalidOperationException($"Failed to select first point at ({firstXmm}, {firstYmm})");

    // Select second point
    bool secondSelected = _activeModel.Extension.SelectByID2(
        string.Empty, "SKETCHPOINT", 
        secondXmm * MmToMeters, secondYmm * MmToMeters, 0, 
        true, 0, null, 0);
        
    if (!secondSelected)
    {
        _activeModel.ClearSelection2(true);
        throw new InvalidOperationException($"Failed to select second point at ({secondXmm}, {secondYmm})");
    }

    // Create dimension
    var displayDim = (DisplayDimension)_activeModel.AddDimension2(
        textXmm * MmToMeters, textYmm * MmToMeters, 0);
        
    _activeModel.ClearSelection2(true);

    if (displayDim == null)
        throw new InvalidOperationException("Failed to create dimension");

    if (valueMm.HasValue)
    {
        var swDim = displayDim.GetDimension();
        swDim.SystemValue = valueMm.Value * MmToMeters;
    }
}
        public void AddDimension(double pickXmm, double pickYmm, double textXmm, double textYmm, double? valueMm = null)
{
    EnsureSketchOpen();
    _activeModel.ClearSelection2(true);

    // Select entity at pick point
    bool selected = _activeModel.Extension.SelectByID2(
        string.Empty, "SKETCHSEGMENT",
        pickXmm * MmToMeters, pickYmm * MmToMeters, 0,
        false, 0, null, 0);

    if (!selected)
        throw new InvalidOperationException($"Failed to select entity at ({pickXmm}, {pickYmm})");

    // Create dimension
    var displayDim = (DisplayDimension)_activeModel.AddDimension2(
        textXmm * MmToMeters, textYmm * MmToMeters, 0);

    _activeModel.ClearSelection2(true);

    if (displayDim == null)
        throw new InvalidOperationException("Failed to create dimension");

    if (valueMm.HasValue)
    {
        var swDim = displayDim.GetDimension();
        swDim.SystemValue = valueMm.Value * MmToMeters;
    }
}

    public void EndSketch()
    {
        EnsureSketchOpen();
        _sketchManager.InsertSketch(true);
        _isSketchOpen = false;
    }

    public void Extrude(double depthMm, bool midPlane = false, bool isCut = false)
    {
        EnsureModelReady();
        
        if (_isSketchOpen)
            EndSketch();

        int endCondition = midPlane
            ? (int)swEndConditions_e.swEndCondMidPlane
            : (int)swEndConditions_e.swEndCondBlind;

        var result = _activeModel.FeatureManager.FeatureExtrusion2(
            !isCut, false, false,
            endCondition, 0,
            depthMm * MmToMeters, 0,
            false, false, false, false,
            0, 0, false, false, false, false,
            true, true, true, 0, 0, false);

        if (result == null)
            throw new InvalidOperationException("Failed to create extrusion");
    }

    public void Rebuild()
    {
        EnsureModelReady();
        _activeModel.ForceRebuild3(false);
    }

    public string SavePart()
    {
        EnsureModelReady();
        
        if (string.IsNullOrWhiteSpace(_activePartName) || string.IsNullOrWhiteSpace(_outputFolder))
            throw new InvalidOperationException("No active part context. Use BeginPartSketch first.");

        string fullPath = Path.Combine(_outputFolder, _activePartName + ".SLDPRT");
        _activeModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 
                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        return fullPath;
    }

    public string SaveAndClose(bool closeDocument)
    {
        string savedPath = SavePart();
        
        if (closeDocument)
        {
            string title = _activeModel.GetTitle();
            _swApp.CloseDoc(title);
            _activeModel = null;
            _sketchManager = null;
            _isSketchOpen = false;
        }

        return savedPath;
    }

    public ModelDoc2 GetActiveModel()
    {
        EnsureModelReady();
        return _activeModel;
    }

    // Add these enum definitions if they're not already present at the top of your file
// ============================================================================
// FILLET METHODS
// ============================================================================

/// <summary>
/// Create sketch fillet using one or two selected entities.
/// </summary>
/// <param name="radiusMm">Fillet radius in millimeters</param>
/// <param name="firstEntity">First entity reference (segment or point near corner)</param>
/// <param name="secondEntity">Optional second entity reference</param>
/// <param name="constrainedCorners">True to create constraints, false for unconstrained fillet</param>
public void CreateSketchFillet(
    double radiusMm,
    SketchEntityReference firstEntity,
    SketchEntityReference secondEntity = default,
    bool constrainedCorners = true)
{
    EnsureSketchOpen();
    
    if (!firstEntity.IsValid())
        throw new ArgumentException("First entity is required for sketch fillet", nameof(firstEntity));

    _activeModel.ClearSelection2(true);

    // Select first entity
    bool firstSelected = _activeModel.Extension.SelectByID2(
        firstEntity.Name ?? string.Empty,
        firstEntity.Type,
        firstEntity.XMm * MmToMeters,
        firstEntity.YMm * MmToMeters,
        firstEntity.ZMm * MmToMeters,
        false, 0, null, 0);

    if (!firstSelected)
        throw new InvalidOperationException("Failed to select first fillet entity");

    // Select second entity if provided
    if (secondEntity.IsValid())
    {
        bool secondSelected = _activeModel.Extension.SelectByID2(
            secondEntity.Name ?? string.Empty,
            secondEntity.Type,
            secondEntity.XMm * MmToMeters,
            secondEntity.YMm * MmToMeters,
            secondEntity.ZMm * MmToMeters,
            true, 0, null, 0);

        if (!secondSelected)
        {
            _activeModel.ClearSelection2(true);
            throw new InvalidOperationException("Failed to select second fillet entity");
        }
    }

    // Create fillet (1 = constrained corners, 0 = unconstrained)
    _sketchManager.CreateFillet(radiusMm * MmToMeters, constrainedCorners ? 1 : 0);
    _activeModel.ClearSelection2(true);
}

        /// <summary>
        /// Create sketch fillet by picking a point near the corner to fillet.
        /// </summary>
        /// <param name="radiusMm">Fillet radius in millimeters</param>
        /// <param name="cornerPoint">Point reference near the corner to fillet</param>
        /// <param name="constrainedCorners">True to create constraints, false for unconstrained fillet</param>
        public void CreateSketchFillet(
            double radiusMm,
            SketchEntityReference cornerPoint,
            bool constrainedCorners = true)
        {
            CreateSketchFillet(radiusMm, cornerPoint, default, constrainedCorners);
        }

// ============================================================================
// CHAMFER METHODS
// ============================================================================

/// <summary>
/// Create sketch chamfer using one or two selected entities.
/// </summary>
/// <param name="chamferMode">Distance-Angle or Distance-Distance mode</param>
/// <param name="distance1Mm">First distance in millimeters</param>
/// <param name="distance2OrAngle">Second distance (mm) or angle (degrees) depending on mode</param>
/// <param name="firstEntity">First entity reference</param>
/// <param name="secondEntity">Optional second entity reference</param>
        public void CreateSketchChamfer(
            SketchChamferMode chamferMode,
            double distance1Mm,
            double distance2OrAngle,
            SketchEntityReference firstEntity,
            SketchEntityReference secondEntity = default)
        {
            EnsureSketchOpen();

            if (!firstEntity.IsValid())
                throw new ArgumentException("First entity is required for sketch chamfer", nameof(firstEntity));

            _activeModel.ClearSelection2(true);

            // Select first entity
            bool firstSelected = _activeModel.Extension.SelectByID2(
                firstEntity.Name ?? string.Empty,
                firstEntity.Type,
                firstEntity.XMm * MmToMeters,
                firstEntity.YMm * MmToMeters,
                firstEntity.ZMm * MmToMeters,
                false, 0, null, 0);

            if (!firstSelected)
                throw new InvalidOperationException("Failed to select first chamfer entity");

            // Select second entity if provided
            if (secondEntity.IsValid())
            {
                bool secondSelected = _activeModel.Extension.SelectByID2(
                    secondEntity.Name ?? string.Empty,
                    secondEntity.Type,
                    secondEntity.XMm * MmToMeters,
                    secondEntity.YMm * MmToMeters,
                    secondEntity.ZMm * MmToMeters,
                    true, 0, null, 0);

                if (!secondSelected)
                {
                    _activeModel.ClearSelection2(true);
                    throw new InvalidOperationException("Failed to select second chamfer entity");
                }
            }

            // Set chamfer type
            int type = chamferMode == SketchChamferMode.DistanceDistance
                ? (int)swSketchChamferType_e.swSketchChamfer_DistanceDistance
                : (int)swSketchChamferType_e.swSketchChamfer_DistanceAngle;

            double value1 = distance1Mm * MmToMeters;
            double value2;

            if (chamferMode == SketchChamferMode.DistanceDistance)
            {
                // If second distance is omitted or NaN, use equal-distance chamfer
                double distance2Mm = double.IsNaN(distance2OrAngle) ? distance1Mm : distance2OrAngle;
                value2 = distance2Mm * MmToMeters;
            }
            else
            {
                // Distance-Angle mode: angle is in degrees (convert to radians for SOLIDWORKS)
                if (double.IsNaN(distance2OrAngle))
                    throw new ArgumentException("Distance-Angle chamfer requires angle value", nameof(distance2OrAngle));
                
                value2 = distance2OrAngle * Math.PI / 180.0; // Degrees to radians
            }

            _sketchManager.CreateChamfer(type, value1, value2);
            _activeModel.ClearSelection2(true);
        }

        /// <summary>
        /// Create sketch chamfer by picking a point near the corner.
        /// </summary>
        /// <param name="chamferMode">Distance-Angle or Distance-Distance mode</param>
        /// <param name="distance1Mm">First distance in millimeters</param>
        /// <param name="distance2OrAngle">Second distance (mm) or angle (degrees) depending on mode</param>
        /// <param name="cornerPoint">Point reference near the corner to chamfer</param>
        public void CreateSketchChamfer(
            SketchChamferMode chamferMode,
            double distance1Mm,
            double distance2OrAngle,
            SketchEntityReference cornerPoint)
        {
            CreateSketchChamfer(chamferMode, distance1Mm, distance2OrAngle, cornerPoint, default);
        }

        /// <summary>
        /// Create equal-distance sketch chamfer (Distance-Distance mode with equal distances).
        /// </summary>
        /// <param name="distanceMm">Equal chamfer distance in millimeters</param>
        /// <param name="firstEntity">First entity reference</param>
        /// <param name="secondEntity">Optional second entity reference</param>
        public void CreateSketchChamfer(
            double distanceMm,
            SketchEntityReference firstEntity,
            SketchEntityReference secondEntity = default)
        {
            CreateSketchChamfer(SketchChamferMode.DistanceDistance, distanceMm, distanceMm, firstEntity, secondEntity);
        }

        /// <summary>
        /// Create equal-distance sketch chamfer by picking a point near the corner.
        /// </summary>
        /// <param name="distanceMm">Equal chamfer distance in millimeters</param>
        /// <param name="cornerPoint">Point reference near the corner to chamfer</param>
        public void CreateSketchChamfer(
            double distanceMm,
            SketchEntityReference cornerPoint)
        {
            CreateSketchChamfer(SketchChamferMode.DistanceDistance, distanceMm, distanceMm, cornerPoint, default);
        }

    // Private helpers
    private void StartSketchOnPlane(string planeName)
    {
        _activeModel.ClearSelection2(true);
        
        bool selected = _activeModel.Extension.SelectByID2(planeName, PlaneType, 0, 0, 0, false, 0, null, 0);
        
        if (!selected)
            throw new InvalidOperationException($"Failed to select plane: {planeName}");

        _sketchManager.InsertSketch(true);
        _isSketchOpen = true;
    }

    private void CloseOpenSketch()
    {
        if (_isSketchOpen)
        {
            _sketchManager.InsertSketch(true);
            _isSketchOpen = false;
        }
    }

    private static string GetPlaneName(SketchPlaneName0 sketchPlane)
{
    return sketchPlane switch
    {
        SketchPlaneName0.Top => "Ebene oben",
        SketchPlaneName0.Right or SketchPlaneName0.Side => "Ebene rechts",
        _ => "Ebene vorne"  // Front is default
    };
}

    private static string GetConstraintToken(SketchRelationType relationType)
    {
        return relationType switch
        {
            SketchRelationType.Horizontal => "sgHORIZONTAL",
            SketchRelationType.Vertical => "sgVERTICAL",
            SketchRelationType.Parallel => "sgPARALLEL",
            SketchRelationType.Perpendicular => "sgPERPENDICULAR",
            SketchRelationType.Tangent => "sgTANGENT",
            SketchRelationType.Equal => "sgEQUAL",
            SketchRelationType.Symmetric => "sgSYMMETRIC",
            SketchRelationType.Coincident => "sgCOINCIDENT",
            SketchRelationType.Collinear => "sgCOLINEAR",
            _ => throw new ArgumentOutOfRangeException(nameof(relationType), relationType, "Unsupported relation type")
        };
    }

    private void EnsureModelReady()
    {
        if (_activeModel == null)
            throw new InvalidOperationException("No active model. Call BeginPartSketch or BeginSketch first.");
    }

    private void EnsureSketchOpen()
    {
        EnsureModelReady();
        
        if (!_isSketchOpen)
            throw new InvalidOperationException("Sketch is not open. Call BeginSketch first.");
    }
    public void ClearSelection()
    {
        EnsureModelReady();
        _activeModel.ClearSelection2(true);
    }
    public bool IsSketchActive()
    {
        if (_activeModel == null) return false;
        return _sketchManager.ActiveSketch != null;
    }
        public object GetActiveSketch()
    {
        EnsureSketchOpen();
        return _sketchManager.ActiveSketch;
    }
    public void DisableSketchInference()
{
    // Disable sketch snapping/inference
    _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
}

public void EnableSketchInference()
{
    // Re-enable sketch snapping/inference
    _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
}
    
}