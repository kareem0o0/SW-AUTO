using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Creates the NDE stator press ring.
///
/// This is one of the more complex parts in the project.
/// In simple terms, it is:
/// - a thick outer ring
/// - a thinner inner press ring region
/// - chamfer detail
/// - one pocket feature
/// - a circular pattern of that pocket
/// </summary>
public sealed class StatorPressringNdePart
{
    
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    public StatorPressringNdePart(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        _pdm = pdm ?? throw new ArgumentNullException(nameof(pdm));
    }

    // File and save settings.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "StatorPressringNDE.SLDPRT";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    // Main editable geometry values.
    public double OuterDiameter { get; set; } = 1.1;
    public double InnerDiameter { get; set; } = 0.84;
    public double PressRingOuterDiameter { get; set; } = 0.86;
    public double RingThickness { get; set; } = 0.028;
    public double PressRingThickness { get; set; } = 0.002;
    public double BaseInnerChamferDistance { get; set; } = 0.02;
    public double BaseInnerChamferAngleDeg { get; set; } = 30.0;
    public double PocketCenterRadius { get; set; } = 0.52;
    public double PocketWidth { get; set; } = 0.043;
    public double PocketHeight { get; set; } = 0.036;
    public double PocketCornerRadius { get; set; } = 0.005;
    public int PocketCount { get; set; } = 8;
    public string MaterialName { get; set; } = "AISI 1020";

    private string GetRequiredOutputFolder() => AutomationSupport.RequireText(OutputFolder, nameof(OutputFolder), nameof(StatorPressringNdePart));
    private string GetRequiredLocalFileName() => AutomationSupport.RequireText(LocalFileName, nameof(LocalFileName), nameof(StatorPressringNdePart));
    private AutomationUiScope BeginAutomationUiSuppression() => new(_swApp);

    /// <summary>
    /// Creates the stator press ring NDE model and saves it.
    /// </summary>
    public string Create()
    {
        
        using var automationUi = BeginAutomationUiSuppression();

        // Read the current object values first.
        string outFolder = GetRequiredOutputFolder();
        bool closeAfterCreate = CloseAfterCreate;
        bool saveToPdm = SaveToPdm;

        bool SelectSketchByIndex(ModelDoc2 model, int index)
        {
            return model.Extension.SelectByID2($"Skizze{index}", "SKETCH", 0, 0, 0, false, 0, null, 0)
                || model.Extension.SelectByID2($"Sketch{index}", "SKETCH", 0, 0, 0, false, 0, null, 0);
        }

        // Main dimensions (m) - change these only.
        double outerDiameter = OuterDiameter;
        double innerDiameter = InnerDiameter;
        double pressRingOuterDiameter = PressRingOuterDiameter;
        double ringThicknessValue = RingThickness;
        double pressRingThicknessValue = PressRingThickness;
        double baseInnerChamferDistanceValue = BaseInnerChamferDistance;
        double baseInnerChamferAngleDeg = BaseInnerChamferAngleDeg;
        double pocketCenterRadiusValue = PocketCenterRadius;
        double pocketWidthValue = PocketWidth;
        double pocketHeightValue = PocketHeight;
        double pocketCornerRadiusValue = PocketCornerRadius;
        int pocketCount = PocketCount;
        string materialName = MaterialName;

        // Derived dimensions
        double outerRadius = outerDiameter / 2.0;
        double innerRadius = innerDiameter / 2.0;
        double pressRingOuterRadius = pressRingOuterDiameter / 2.0;
        double ringThickness = ringThicknessValue;
        double pressRingThickness = pressRingThicknessValue;
        double baseInnerChamferDistance = baseInnerChamferDistanceValue;
        double baseInnerChamferAngleRadians = baseInnerChamferAngleDeg * Math.PI / 180.0;
        double pocketCenterY = pocketCenterRadiusValue;
        double pocketCornerRadius = pocketCornerRadiusValue;
        double pocketHalfWidth = pocketWidthValue / 2.0;
        double pocketHalfHeight = pocketHeightValue / 2.0;
        double pocketCircleTopCenterY = pocketCenterY + pocketHalfHeight;
        double pocketCircleBottomCenterY = pocketCenterY - pocketHalfHeight;
        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        try
        {
            Dimension swDim = null;
            DisplayDimension displayDim = null;
            bool sketchInferenceWasEnabled = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference);

            // Local helper:
            // creates the pocket cut from the already prepared pocket sketch.
            Feature CreatePocketCut(bool reverseDirection)
            {
                swModel.ClearSelection2(true);
                if (!SelectSketchByIndex(swModel, 2))
                    throw new Exception("Could not select stator pressring NDE pocket sketch");

                return swModel.FeatureManager.FeatureCut4(
                    true, false, reverseDirection,
                    (int)swEndConditions_e.swEndCondThroughAll, 0,
                    0, 0,
                    false, false, false, false,
                    0, 0, false, false, false, false,
                    false, true, true, true, true, false,
                    0, 0, false, false);
            }

            // Local helper:
            // converts diagonal reference lines inside the pocket sketch into construction geometry
            // so they assist dimensioning without affecting the final cut profile.
            void ConvertPocketReferenceDiagonalsToConstruction()
            {
                Sketch activePocketSketch = swModel.GetActiveSketch2() as Sketch;
                if (activePocketSketch == null)
                    throw new Exception("Could not access active stator pressring NDE pocket sketch");

                object[] sketchSegments = activePocketSketch.GetSketchSegments() as object[];
                if (sketchSegments == null || sketchSegments.Length == 0)
                    throw new Exception("Could not access stator pressring NDE pocket sketch segments");

                const double axisTolerance = 1e-9;
                var diagonalSegments = new System.Collections.Generic.List<SketchSegment>();

                foreach (object segmentObj in sketchSegments)
                {
                    SketchLine sketchLine = segmentObj as SketchLine;
                    if (sketchLine == null)
                        continue;

                    SketchPoint startPoint = sketchLine.GetStartPoint2();
                    SketchPoint endPoint = sketchLine.GetEndPoint2();
                    if (startPoint == null || endPoint == null)
                        continue;

                    double dx = Math.Abs(endPoint.X - startPoint.X);
                    double dy = Math.Abs(endPoint.Y - startPoint.Y);
                    if (dx > axisTolerance && dy > axisTolerance)
                        diagonalSegments.Add((SketchSegment)sketchLine);
                }

                if (diagonalSegments.Count == 0)
                    return;

                for (int i = 0; i < diagonalSegments.Count; i++)
                {
                    diagonalSegments[i].ConstructionGeometry = true;
                }

                swModel.ClearSelection2(true);
            }

            // Local helper:
            // returns the midpoint of a sketch segment so nearby reference lines can be found.
            double[] GetSketchSegmentMidpoint(SketchSegment segment)
            {
                SketchLine line = segment as SketchLine;
                if (line == null)
                    throw new Exception("Expected stator pressring NDE pocket reference line");

                SketchPoint startPoint = line.GetStartPoint2();
                SketchPoint endPoint = line.GetEndPoint2();
                if (startPoint == null || endPoint == null)
                    throw new Exception("Could not access stator pressring NDE pocket reference endpoints");

                return new double[]
                {
                    (startPoint.X + endPoint.X) / 2.0,
                    (startPoint.Y + endPoint.Y) / 2.0
                };
            }

            // Local helper:
            // finds the sketch segment closest to an expected location.
            SketchSegment FindClosestPocketReferenceSegment(object[] sketchSegments, double targetX, double targetY, string label)
            {
                SketchSegment bestSegment = null;
                double bestDistance = double.MaxValue;

                foreach (object segmentObj in sketchSegments)
                {
                    SketchSegment candidate = segmentObj as SketchSegment;
                    if (candidate == null)
                        continue;

                    double[] midpoint = GetSketchSegmentMidpoint(candidate);
                    double distance = Math.Sqrt(
                        Math.Pow(midpoint[0] - targetX, 2)
                        + Math.Pow(midpoint[1] - targetY, 2));

                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        bestSegment = candidate;
                    }
                }

                if (bestSegment == null)
                    throw new Exception($"Could not find stator pressring NDE pocket reference {label}");

                return bestSegment;
            }

            SketchPoint GetLeftEndpoint(SketchSegment segment)
            {
                SketchLine line = segment as SketchLine;
                if (line == null)
                    throw new Exception("Expected stator pressring NDE pocket reference line");

                SketchPoint startPoint = line.GetStartPoint2();
                SketchPoint endPoint = line.GetEndPoint2();
                if (startPoint == null || endPoint == null)
                    throw new Exception("Could not access stator pressring NDE pocket reference endpoints");

                return startPoint.X <= endPoint.X ? startPoint : endPoint;
            }

            SketchPoint GetRightEndpoint(SketchSegment segment)
            {
                SketchLine line = segment as SketchLine;
                if (line == null)
                    throw new Exception("Expected stator pressring NDE pocket reference line");

                SketchPoint startPoint = line.GetStartPoint2();
                SketchPoint endPoint = line.GetEndPoint2();
                if (startPoint == null || endPoint == null)
                    throw new Exception("Could not access stator pressring NDE pocket reference endpoints");

                return startPoint.X >= endPoint.X ? startPoint : endPoint;
            }

            Edge FindBaseInnerChamferEdge(Feature baseFeature)
            {
                object[] baseFaces = baseFeature?.GetFaces() as object[];
                if (baseFaces == null || baseFaces.Length == 0)
                    throw new Exception("Could not access stator pressring NDE base ring faces for chamfer");

                double faceTolerance = 0.00001;
                double radiusTolerance = 0.00005;
                Face2 bottomPlanarFace = null;

                foreach (object faceObj in baseFaces)
                {
                    Face2 candidateFace = faceObj as Face2;
                    if (candidateFace == null)
                        continue;

                    Surface candidateSurface = candidateFace.GetSurface();
                    if (candidateSurface == null || !candidateSurface.IsPlane())
                        continue;

                    double[] candidateBox = candidateFace.GetBox() as double[];
                    if (candidateBox == null || candidateBox.Length < 6)
                        continue;

                    if (Math.Abs(candidateBox[2]) < faceTolerance && Math.Abs(candidateBox[5]) < faceTolerance)
                    {
                        bottomPlanarFace = candidateFace;
                        break;
                    }
                }

                if (bottomPlanarFace == null)
                    throw new Exception("Could not find stator pressring NDE base bottom face for chamfer");

                object[] faceEdges = bottomPlanarFace.GetEdges() as object[];
                if (faceEdges == null || faceEdges.Length == 0)
                    throw new Exception("Could not access stator pressring NDE base bottom-face edges for chamfer");

                foreach (object edgeObj in faceEdges)
                {
                    Edge candidateEdge = edgeObj as Edge;
                    if (candidateEdge == null)
                        continue;

                    Curve candidateCurve = candidateEdge.GetCurve();
                    if (candidateCurve == null || !candidateCurve.IsCircle())
                        continue;

                    double[] circleParams = candidateCurve.CircleParams as double[];
                    if (circleParams == null || circleParams.Length < 7)
                        continue;

                    double candidateRadius = circleParams[6];
                    if (Math.Abs(candidateRadius - innerRadius) < radiusTolerance)
                        return candidateEdge;
                }

                throw new Exception("Could not find stator pressring NDE base inner circular edge for chamfer");
            }

            if (!Directory.Exists(outFolder))
                Directory.CreateDirectory(outFolder);

            // Phase 1:
            // Create the main ring body from the default part template.
            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
            if (swModel == null)
                throw new Exception("Failed to create new part");

            swSketchManager = swModel.SketchManager;

            PartDoc pressRingPart = swModel as PartDoc;
            pressRingPart?.SetMaterialPropertyName2("", "", Name: materialName);

            bool selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select Front Plane for stator pressring NDE");

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCircleByRadius(0, 0, 0, outerRadius);
            swSketchManager.CreateCircleByRadius(0, 0, 0, innerRadius);

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", outerRadius, 0, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE outer circle");
            displayDim = (DisplayDimension)swModel.AddDimension2(outerRadius + 0.02, 0.02, 0);
            if (displayDim == null)
                throw new Exception("Could not create stator pressring NDE outer diameter dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception("Could not access stator pressring NDE outer diameter dimension");
            swDim.SystemValue = outerDiameter;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", innerRadius, 0, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE inner circle");
            displayDim = (DisplayDimension)swModel.AddDimension2(innerRadius + 0.02, 0.02, 0);
            if (displayDim == null)
                throw new Exception("Could not create stator pressring NDE inner diameter dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception("Could not access stator pressring NDE inner diameter dimension");
            swDim.SystemValue = innerDiameter;

            swModel.ClearSelection2(true);
            if (!SelectSketchByIndex(swModel, 1))
                throw new Exception("Could not select stator pressring NDE base sketch");
            Feature baseRingFeature = swModel.FeatureManager.FeatureExtrusion2(
                true, false, false,
                (int)swEndConditions_e.swEndCondBlind, 0,
                ringThickness, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                true, true, true, 0, 0, false);
            if (baseRingFeature == null)
                throw new Exception("Failed to create stator pressring NDE base ring");

            swModel.ClearSelection2(true);
            Edge baseInnerChamferEdge = FindBaseInnerChamferEdge(baseRingFeature);
            Entity baseInnerChamferEdgeEntity = baseInnerChamferEdge as Entity;
            if (baseInnerChamferEdgeEntity == null || !baseInnerChamferEdgeEntity.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE base inner edge for chamfer");

            Feature baseInnerChamferFeature = swModel.FeatureManager.InsertFeatureChamfer(
                0,
                (int)swChamferType_e.swChamferAngleDistance,
                baseInnerChamferDistance,
                baseInnerChamferAngleRadians,
                0,
                0,
                0,
                0);
            if (baseInnerChamferFeature == null)
            {
                swModel.ClearSelection2(true);
                if (!baseInnerChamferEdgeEntity.Select4(false, null))
                    throw new Exception("Could not reselect stator pressring NDE base inner edge for flipped chamfer");

                baseInnerChamferFeature = swModel.FeatureManager.InsertFeatureChamfer(
                    (int)swFeatureChamferOption_e.swFeatureChamferFlipDirection,
                    (int)swChamferType_e.swChamferAngleDistance,
                    baseInnerChamferDistance,
                    baseInnerChamferAngleRadians,
                    0,
                    0,
                    0,
                    0);
            }
            if (baseInnerChamferFeature == null)
                throw new Exception("Failed to create stator pressring NDE base inner chamfer");

            // Phase 2:
            // Sketch one pocket reference rectangle on the top face.
            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2(
                "",
                "FACE",
                (outerRadius + innerRadius) / 2.0,
                0,
                ringThickness,
                false,
                0,
                null,
                0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE top face for pocket");

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCenterRectangle(0, pocketCenterY, 0, pocketHalfWidth, pocketCircleTopCenterY, 0);
            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, pocketCircleTopCenterY, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE pocket reference top edge");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(0, pocketCircleTopCenterY + 0.015, 0);
            if (displayDim == null)
                throw new Exception("Could not create stator pressring NDE pocket width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception("Could not access stator pressring NDE pocket width dimension");
            swDim.SystemValue = pocketHalfWidth * 2.0;
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", pocketHalfWidth, pocketCircleBottomCenterY, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE pocket reference bottom-right point");
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", pocketHalfWidth, pocketCircleTopCenterY, 0, true, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE pocket reference top-right point");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(pocketHalfWidth + 0.025, pocketCenterY, 0);
            if (displayDim == null)
                throw new Exception("Could not create stator pressring NDE pocket height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception("Could not access stator pressring NDE pocket height dimension");
            swDim.SystemValue = pocketHalfHeight * 2.0;
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", 0, pocketCenterY, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE pocket reference center point");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select sketch origin for stator pressring NDE pocket alignment");
            swModel.SketchAddConstraints("sgVERTICALPOINTS2D");
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", 0, pocketCenterY, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not reselect stator pressring NDE pocket reference center point");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select sketch origin for stator pressring NDE pocket location");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(-0.04, pocketCenterY / 2.0, 0);
            if (displayDim == null)
                throw new Exception("Could not create stator pressring NDE pocket center radius dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception("Could not access stator pressring NDE pocket center radius dimension");
            swDim.SystemValue = pocketCenterY;
            swModel.ClearSelection2(true);

            ConvertPocketReferenceDiagonalsToConstruction();

            Sketch activePocketReferenceSketch = swModel.GetActiveSketch2() as Sketch;
            if (activePocketReferenceSketch == null)
                throw new Exception("Could not access stator pressring NDE pocket reference sketch");
            object[] pocketReferenceSegments = activePocketReferenceSketch.GetSketchSegments() as object[];
            if (pocketReferenceSegments == null || pocketReferenceSegments.Length < 4)
                throw new Exception("Could not access stator pressring NDE pocket reference segments");

            // Convert the simple reference rectangle into the final rounded pocket profile.
            SketchSegment pocketReferenceTopEdge = FindClosestPocketReferenceSegment(
                pocketReferenceSegments, 0, pocketCircleTopCenterY, "top edge");
            SketchSegment pocketReferenceRightEdge = FindClosestPocketReferenceSegment(
                pocketReferenceSegments, pocketHalfWidth, pocketCenterY, "right edge");
            SketchSegment pocketReferenceBottomEdge = FindClosestPocketReferenceSegment(
                pocketReferenceSegments, 0, pocketCircleBottomCenterY, "bottom edge");
            SketchSegment pocketReferenceLeftEdge = FindClosestPocketReferenceSegment(
                pocketReferenceSegments, -pocketHalfWidth, pocketCenterY, "left edge");

            SketchPoint topLeftReferenceCorner = GetLeftEndpoint(pocketReferenceTopEdge);
            SketchPoint topRightReferenceCorner = GetRightEndpoint(pocketReferenceTopEdge);
            SketchPoint bottomLeftReferenceCorner = GetLeftEndpoint(pocketReferenceBottomEdge);
            SketchPoint bottomRightReferenceCorner = GetRightEndpoint(pocketReferenceBottomEdge);

            swModel.ClearSelection2(true);
            if (!pocketReferenceTopEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE pocket reference top edge for horizontal relation");
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            if (!pocketReferenceBottomEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE pocket reference bottom edge for horizontal relation");
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            if (!pocketReferenceRightEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE pocket reference right edge for vertical relation");
            swModel.SketchAddConstraints("sgVERTICAL");
            swModel.ClearSelection2(true);

            if (!pocketReferenceLeftEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE pocket reference left edge for vertical relation");
            swModel.SketchAddConstraints("sgVERTICAL");
            swModel.ClearSelection2(true);

            if (!pocketReferenceTopEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE pocket reference top edge for construction");
            if (!pocketReferenceRightEdge.Select4(true, null))
                throw new Exception("Could not select stator pressring NDE pocket reference right edge for construction");
            if (!pocketReferenceBottomEdge.Select4(true, null))
                throw new Exception("Could not select stator pressring NDE pocket reference bottom edge for construction");
            if (!pocketReferenceLeftEdge.Select4(true, null))
                throw new Exception("Could not select stator pressring NDE pocket reference left edge for construction");
            swSketchManager.CreateConstructionGeometry();
            swModel.ClearSelection2(true);

            SketchSegment topLeftArcSegment = swSketchManager.CreateArc(
                -pocketHalfWidth, pocketCircleTopCenterY, 0,
                -pocketHalfWidth + pocketCornerRadius, pocketCircleTopCenterY, 0,
                -pocketHalfWidth, pocketCircleTopCenterY - pocketCornerRadius, 0,
                1) as SketchSegment;
            SketchSegment topRightArcSegment = swSketchManager.CreateArc(
                pocketHalfWidth, pocketCircleTopCenterY, 0,
                pocketHalfWidth - pocketCornerRadius, pocketCircleTopCenterY, 0,
                pocketHalfWidth, pocketCircleTopCenterY - pocketCornerRadius, 0,
                -1) as SketchSegment;
            SketchSegment bottomRightArcSegment = swSketchManager.CreateArc(
                pocketHalfWidth, pocketCircleBottomCenterY, 0,
                pocketHalfWidth, pocketCircleBottomCenterY + pocketCornerRadius, 0,
                pocketHalfWidth - pocketCornerRadius, pocketCircleBottomCenterY, 0,
                -1) as SketchSegment;
            SketchSegment bottomLeftArcSegment = swSketchManager.CreateArc(
                -pocketHalfWidth, pocketCircleBottomCenterY, 0,
                -pocketHalfWidth + pocketCornerRadius, pocketCircleBottomCenterY, 0,
                -pocketHalfWidth, pocketCircleBottomCenterY + pocketCornerRadius, 0,
                -1) as SketchSegment;

            if (topLeftArcSegment == null || topRightArcSegment == null
                || bottomRightArcSegment == null || bottomLeftArcSegment == null)
            {
                throw new Exception("Could not create stator pressring NDE pocket corner arcs");
            }

            SketchArc topLeftArc = topLeftArcSegment as SketchArc;
            SketchArc topRightArc = topRightArcSegment as SketchArc;
            SketchArc bottomRightArc = bottomRightArcSegment as SketchArc;
            SketchArc bottomLeftArc = bottomLeftArcSegment as SketchArc;
            if (topLeftArc == null || topRightArc == null || bottomRightArc == null || bottomLeftArc == null)
                throw new Exception("Could not access stator pressring NDE pocket arc geometry");

            SketchPoint topLeftArcCenter = topLeftArc.GetCenterPoint2();
            SketchPoint topRightArcCenter = topRightArc.GetCenterPoint2();
            SketchPoint bottomRightArcCenter = bottomRightArc.GetCenterPoint2();
            SketchPoint bottomLeftArcCenter = bottomLeftArc.GetCenterPoint2();
            if (topLeftArcCenter == null || topRightArcCenter == null
                || bottomRightArcCenter == null || bottomLeftArcCenter == null)
                throw new Exception("Could not access stator pressring NDE pocket arc centers");

            swModel.ClearSelection2(true);
            if (!topLeftArcCenter.Select4(false, null) || !topLeftReferenceCorner.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE top-left arc center");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!topRightArcCenter.Select4(false, null) || !topRightReferenceCorner.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE top-right arc center");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!bottomRightArcCenter.Select4(false, null) || !bottomRightReferenceCorner.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE bottom-right arc center");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!bottomLeftArcCenter.Select4(false, null) || !bottomLeftReferenceCorner.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE bottom-left arc center");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!topLeftArcSegment.Select4(false, null)
                || !topRightArcSegment.Select4(true, null)
                || !bottomRightArcSegment.Select4(true, null)
                || !bottomLeftArcSegment.Select4(true, null))
            {
                throw new Exception("Could not select stator pressring NDE pocket arcs for equal relation");
            }
            swModel.SketchAddConstraints("sgSAMELENGTH");
            swModel.ClearSelection2(true);
            SketchSegment topEdge = swSketchManager.CreateLine(
                -pocketHalfWidth + pocketCornerRadius, pocketCircleTopCenterY, 0,
                pocketHalfWidth - pocketCornerRadius, pocketCircleTopCenterY, 0) as SketchSegment;
            SketchSegment rightEdge = swSketchManager.CreateLine(
                pocketHalfWidth, pocketCircleTopCenterY - pocketCornerRadius, 0,
                pocketHalfWidth, pocketCircleBottomCenterY + pocketCornerRadius, 0) as SketchSegment;
            SketchSegment bottomEdge = swSketchManager.CreateLine(
                pocketHalfWidth - pocketCornerRadius, pocketCircleBottomCenterY, 0,
                -pocketHalfWidth + pocketCornerRadius, pocketCircleBottomCenterY, 0) as SketchSegment;
            SketchSegment leftEdge = swSketchManager.CreateLine(
                -pocketHalfWidth, pocketCircleBottomCenterY + pocketCornerRadius, 0,
                -pocketHalfWidth, pocketCircleTopCenterY - pocketCornerRadius, 0) as SketchSegment;

            if (topEdge == null || rightEdge == null || bottomEdge == null || leftEdge == null)
            {
                throw new Exception("Could not create stator pressring NDE pocket circle-line profile");
            }

            if (!topEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE top edge");
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            if (!bottomEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE bottom edge");
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            if (!rightEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE right edge");
            swModel.SketchAddConstraints("sgVERTICAL");
            swModel.ClearSelection2(true);

            if (!leftEdge.Select4(false, null))
                throw new Exception("Could not select stator pressring NDE left edge");
            swModel.SketchAddConstraints("sgVERTICAL");
            swModel.ClearSelection2(true);

            if (!topEdge.Select4(false, null) || !pocketReferenceTopEdge.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE top edge to pocket reference top edge");
            swModel.SketchAddConstraints("sgCOLINEAR");
            swModel.ClearSelection2(true);

            if (!bottomEdge.Select4(false, null) || !pocketReferenceBottomEdge.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE bottom edge to pocket reference bottom edge");
            swModel.SketchAddConstraints("sgCOLINEAR");
            swModel.ClearSelection2(true);

            if (!rightEdge.Select4(false, null) || !pocketReferenceRightEdge.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE right edge to pocket reference right edge");
            swModel.SketchAddConstraints("sgCOLINEAR");
            swModel.ClearSelection2(true);

            if (!leftEdge.Select4(false, null) || !pocketReferenceLeftEdge.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE left edge to pocket reference left edge");
            swModel.SketchAddConstraints("sgCOLINEAR");
            swModel.ClearSelection2(true);

            SketchLine topEdgeLine = topEdge as SketchLine;
            SketchLine rightEdgeLine = rightEdge as SketchLine;
            SketchLine bottomEdgeLine = bottomEdge as SketchLine;
            SketchLine leftEdgeLine = leftEdge as SketchLine;
            if (topEdgeLine == null || rightEdgeLine == null || bottomEdgeLine == null || leftEdgeLine == null)
                throw new Exception("Could not access stator pressring NDE pocket line geometry");

            SketchPoint topEdgeStart = topEdgeLine.GetStartPoint2();
            SketchPoint topEdgeEnd = topEdgeLine.GetEndPoint2();
            SketchPoint rightEdgeStart = rightEdgeLine.GetStartPoint2();
            SketchPoint rightEdgeEnd = rightEdgeLine.GetEndPoint2();
            SketchPoint bottomEdgeStart = bottomEdgeLine.GetStartPoint2();
            SketchPoint bottomEdgeEnd = bottomEdgeLine.GetEndPoint2();
            SketchPoint leftEdgeStart = leftEdgeLine.GetStartPoint2();
            SketchPoint leftEdgeEnd = leftEdgeLine.GetEndPoint2();
            if (topEdgeStart == null || topEdgeEnd == null || rightEdgeStart == null || rightEdgeEnd == null
                || bottomEdgeStart == null || bottomEdgeEnd == null || leftEdgeStart == null || leftEdgeEnd == null)
            {
                throw new Exception("Could not access stator pressring NDE pocket line endpoints");
            }

            swModel.ClearSelection2(true);
            if (!topEdgeStart.Select4(false, null) || !topLeftArcSegment.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE top edge start to top-left arc");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!topEdgeEnd.Select4(false, null) || !topRightArcSegment.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE top edge end to top-right arc");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!rightEdgeStart.Select4(false, null) || !topRightArcSegment.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE right edge start to top-right arc");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!rightEdgeEnd.Select4(false, null) || !bottomRightArcSegment.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE right edge end to bottom-right arc");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!bottomEdgeStart.Select4(false, null) || !bottomRightArcSegment.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE bottom edge start to bottom-right arc");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!bottomEdgeEnd.Select4(false, null) || !bottomLeftArcSegment.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE bottom edge end to bottom-left arc");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!leftEdgeStart.Select4(false, null) || !bottomLeftArcSegment.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE left edge start to bottom-left arc");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!leftEdgeEnd.Select4(false, null) || !topLeftArcSegment.Select4(true, null))
                throw new Exception("Could not constrain stator pressring NDE left edge end to top-left arc");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!topLeftArcCenter.Select4(false, null) || !topEdgeStart.Select4(true, null))
                throw new Exception("Could not select stator pressring NDE top-left arc center and top edge start for radius dimension");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(
                -pocketHalfWidth - 0.02,
                pocketCircleTopCenterY + 0.02,
                0);
            if (displayDim == null)
                throw new Exception("Could not create stator pressring NDE pocket arc radius-driving dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception("Could not access stator pressring NDE pocket arc radius-driving dimension");
            swDim.SystemValue = pocketCornerRadius;
            swModel.ClearSelection2(true);

            swSketchManager.InsertSketch(true);

            Feature pocketCutFeature = CreatePocketCut(false);
            if (pocketCutFeature == null)
                pocketCutFeature = CreatePocketCut(true);
            if (pocketCutFeature == null)
                throw new Exception("Failed to create stator pressring NDE pocket cut");

            // Phase 3:
            // Pattern the finished pocket around the ring axis.
            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("Z-Achse", "AXIS", 0, 0, 0, false, 1, null, 0);
            if (!selected)
                throw new Exception("Could not select Z axis for stator pressring NDE pocket pattern");
            if (!pocketCutFeature.Select2(true, 4))
                throw new Exception("Could not select stator pressring NDE pocket cut feature for circular pattern");
            Feature pocketPatternFeature = (Feature)swModel.FeatureManager.FeatureCircularPattern5(
                pocketCount,
                2 * Math.PI,
                false,
                "",
                true,
                true,
                false,
                true,
                false,
                false,
                0,
                0.0,
                "",
                false);
            if (pocketPatternFeature == null)
                throw new Exception("Failed to create stator pressring NDE pocket circular pattern");

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2(
                "",
                "FACE",
                (pressRingOuterRadius + innerRadius) / 2.0,
                0,
                ringThickness,
                false,
                0,
                null,
                0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE top face for inner press ring");
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCircleByRadius(0, 0, 0, pressRingOuterRadius);
            swSketchManager.CreateCircleByRadius(0, 0, 0, innerRadius);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", pressRingOuterRadius, 0, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE inner press ring outer circle");
            displayDim = (DisplayDimension)swModel.AddDimension2(pressRingOuterRadius + 0.02, 0.03, 0);
            if (displayDim == null)
                throw new Exception("Could not create stator pressring NDE inner press ring outer diameter dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception("Could not access stator pressring NDE inner press ring outer diameter dimension");
            swDim.SystemValue = pressRingOuterDiameter;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", innerRadius, 0, 0, false, 0, null, 0);
            if (!selected)
                throw new Exception("Could not select stator pressring NDE inner press ring inner circle");
            displayDim = (DisplayDimension)swModel.AddDimension2(innerRadius + 0.02, -0.03, 0);
            if (displayDim == null)
                throw new Exception("Could not create stator pressring NDE inner press ring inner diameter dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception("Could not access stator pressring NDE inner press ring inner diameter dimension");
            swDim.SystemValue = innerDiameter;

            swSketchManager.InsertSketch(true);
            if (!SelectSketchByIndex(swModel, 3))
                throw new Exception("Could not select stator pressring NDE inner press ring sketch");
            Feature innerPressRingFeature = swModel.FeatureManager.FeatureExtrusion2(
                true, true, false,
                (int)swEndConditions_e.swEndCondBlind, 0,
                pressRingThickness, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                true, true, true, 0, 0, false);
            if (innerPressRingFeature == null)
                throw new Exception("Failed to create stator pressring NDE inner press ring");


            // Save the completed part using the chosen local-or-PDM workflow.
            string savedPath;
            if (saveToPdm)
            {
                savedPath = _pdm.SaveAsPdm(swModel, outFolder, PdmDataCard);
                Console.WriteLine($"Part saved to PDM: {savedPath}");
            }
            else
            {
                savedPath = Path.Combine(outFolder, GetRequiredLocalFileName());
                swModel.SaveAs3(savedPath, 0, 1);
                Console.WriteLine($"Part saved locally: {savedPath}");
            }

            if (closeAfterCreate)
            {
                _swApp.CloseDoc(swModel.GetTitle());
                Console.WriteLine("Part closed after creating.");
            }

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, sketchInferenceWasEnabled);
            Console.WriteLine("Done!");
            return Path.GetFileName(savedPath);
        }
        catch (Exception ex)
        {
            try { _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true); } catch { }
            Console.WriteLine("Fatal error: " + ex);
            return null;
        }
    }

}





