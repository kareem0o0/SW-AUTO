using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Creates the stator distance sheet.
///
/// This part starts from a stator-like annular sheet, then adds the spacer/boss geometry
/// that separates repeated stator packs in the machine stack.
/// </summary>
public sealed class StatorDistanceSheetPart
{
    
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    public StatorDistanceSheetPart(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp;
        _pdm = pdm;
    }

    // File and save settings.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "StatorDistanceBleche.SLDPRT";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    // Main editable geometry values.
    public double OuterDiameter { get; set; } = 0.99;
    public double InnerDiameter { get; set; } = 0.64;
    public double PlateThickness { get; set; } = 0.001;
    public double SlotWidth { get; set; } = 0.0205;
    public double SlotBottomY { get; set; } = 0.32;
    public double SlotTopY { get; set; } = 0.406;
    public double BossRectangleHeight { get; set; } = 0.16;
    public double BossRectangleWidth { get; set; } = 0.008;
    public double BossOuterDiameterOffset { get; set; } = 0.009;
    public double BossCenterlineAngleDeg { get; set; } = 2.96;
    public double BossExtrusionDepth { get; set; } = 0.01;
    public double BossCutOuterTabWidth { get; set; } = 0.002;
    public double BossCutOuterTabHeight { get; set; } = 0.0025;
    public double BossCutTopShelfThickness { get; set; } = 0.0015;
    public double BossCutInnerLegWidth { get; set; } = 0.0015;
    public double BossCutBoundaryExtension { get; set; } = 0.002;
    public int BossCircularPatternCount { get; set; } = 60;
    public int SlotPatternCount { get; set; } = 60;
    public string MaterialName { get; set; } = "AISI 1020";

    private string GetRequiredOutputFolder()
    {
        return OutputFolder;
    }

    private string GetRequiredLocalFileName()
    {
        return LocalFileName;
    }

    private AutomationUiScope BeginAutomationUiSuppression()
    {
        return new AutomationUiScope(_swApp);
    }

    /// <summary>
    /// Creates the stator distance sheet model and saves it.
    /// </summary>
    public string Create()
    {
        
        using var automationUi = BeginAutomationUiSuppression();

        // Read the object's current values first so the rest of the method uses one clear snapshot.
        string outFolder = GetRequiredOutputFolder();
        bool closeAfterCreate = CloseAfterCreate;
        bool saveToPdm = SaveToPdm;
        bool shouldCloseAfterCreate = closeAfterCreate || saveToPdm;
        bool SelectSketchByIndex(ModelDoc2 model, int index)
        {
            return model.Extension.SelectByID2($"Skizze{index}", "SKETCH", 0, 0, 0, false, 0, null, 0)
                || model.Extension.SelectByID2($"Sketch{index}", "SKETCH", 0, 0, 0, false, 0, null, 0);
        }

        // Main dimensions (m) - change these only.
        double outerDiameter = OuterDiameter;
        double innerDiameter = InnerDiameter;
        double plateThicknessValue = PlateThickness;
        double slotWidth = SlotWidth;
        double slotBottomY = SlotBottomY;
        double slotTopY = SlotTopY;
        double bossRectangleHeightValue = BossRectangleHeight;
        double bossRectangleWidthValue = BossRectangleWidth;
        double bossOuterDiameterOffsetValue = BossOuterDiameterOffset;
        double bossCenterlineAngleDeg = BossCenterlineAngleDeg;
        double bossExtrusionDepthValue = BossExtrusionDepth;
        double bossCutOuterTabWidthValue = BossCutOuterTabWidth;
        double bossCutOuterTabHeightValue = BossCutOuterTabHeight;
        double bossCutTopShelfThicknessValue = BossCutTopShelfThickness;
        double bossCutInnerLegWidthValue = BossCutInnerLegWidth;
        double bossCutBoundaryExtensionValue = BossCutBoundaryExtension;
        int bossCircularPatternCount = BossCircularPatternCount;
        int slotPatternCount = SlotPatternCount;
        string materialName = MaterialName;

        // Derived dimensions
        double outerRadius = outerDiameter / 2.0;
        double innerRadius = innerDiameter / 2.0;
        double plateThickness = plateThicknessValue;
        double halfWidth = slotWidth / 2.0;
        double bottomY = slotBottomY;
        double topY = slotTopY;
        double centerY = (topY + bottomY) / 2.0;
        double leftX = -halfWidth;
        double bossRectangleHeight = bossRectangleHeightValue;
        double bossRectangleWidth = bossRectangleWidthValue;
        double bossOuterDiameterOffset = bossOuterDiameterOffsetValue;
        double bossCenterlineAngleRadians = bossCenterlineAngleDeg * Math.PI / 180.0;
        double bossExtrusionDepth = bossExtrusionDepthValue;
        double bossHalfWidth = bossRectangleWidth / 2.0;
        double bossTopCenterRadius = outerRadius - bossOuterDiameterOffset;
        double bossBottomCenterRadius = bossTopCenterRadius - bossRectangleHeight;
        double bossDirectionX = -Math.Sin(bossCenterlineAngleRadians);
        double bossDirectionY = Math.Cos(bossCenterlineAngleRadians);
        double bossNormalX = Math.Cos(bossCenterlineAngleRadians);
        double bossNormalY = Math.Sin(bossCenterlineAngleRadians);
        double bossCutOuterTabWidth = bossCutOuterTabWidthValue;
        double bossCutOuterTabHeight = bossCutOuterTabHeightValue;
        double bossCutTopShelfThickness = bossCutTopShelfThicknessValue;
        double bossCutInnerLegWidth = bossCutInnerLegWidthValue;
        double bossCutBoundaryExtension = bossCutBoundaryExtensionValue;

        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        // Build the distance sheet feature by feature, then save it.
        {
            Dimension swDim = null;
            DisplayDimension displayDim = null;
            SelectionMgr selectionMgr = null;

            if (!Directory.Exists(outFolder))
                Directory.CreateDirectory(outFolder);

            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            if (swModel == null)
                throw new Exception("Failed to create new part");

            swSketchManager = swModel.SketchManager;

            // Apply the requested material name.
            PartDoc statorPart = swModel as PartDoc;
            statorPart.SetMaterialPropertyName2("", "", Name: materialName);

            // Phase 1:
            // Build the annular base ring exactly like a thin stator plate.
            bool selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select Front Plane");

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCircleByRadius(0, 0, 0, outerRadius);
            swSketchManager.CreateCircleByRadius(0, 0, 0, innerRadius);

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", outerRadius, 0, 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(outerRadius + 0.02, 0.02, 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = outerDiameter;

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", innerRadius, 0, 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(innerRadius + 0.02, 0.02, 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = innerDiameter;

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);

            swModel.ClearSelection2(true);
            if (!SelectSketchByIndex(swModel, 1))
                throw new Exception("Could not select base annulus sketch");
            swModel.FeatureManager.FeatureExtrusion2(
                true, false, false,
                (int)swEndConditions_e.swEndCondBlind, 0,
                plateThickness, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                true, true, true, 0, 0, false);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "FACE", (outerRadius + innerRadius) / 2.0, 0, plateThickness, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select top face");

            swSketchManager.InsertSketch(true);

            // Phase 2:
            // Create the main slot profile and dimension it.
            swSketchManager.CreateCenterRectangle(0, centerY, 0, halfWidth, topY, 0);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, topY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot top edge");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(leftX + 0.01, topY + 0.01, 0);
            if (displayDim == null) throw new Exception("Could not create slot width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot width dimension");
            swDim.SystemValue = halfWidth * 2;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom-right point");
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, topY, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot top-right point");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(halfWidth + 0.05, bottomY + 0.05, 0);
            if (displayDim == null) throw new Exception("Could not create slot height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot height dimension");
            swDim.SystemValue = topY - bottomY;

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", 0, centerY, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgVERTICALPOINTS2D");
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom edge");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for slot location");
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY / 2, 0);
            if (displayDim == null) throw new Exception("Could not create slot location dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot location dimension");
            swDim.SystemValue = innerRadius;
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom edge for construction conversion");
            swSketchManager.CreateConstructionGeometry();
            swModel.ClearSelection2(true);

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);

            selectionMgr = swModel.SelectionManager as SelectionMgr;
            if (selectionMgr == null)
                throw new Exception("Could not access selection manager for distance-sheet slot references");

            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", -halfWidth, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select left bottom slot corner");
            SketchPoint leftBottomSlotCorner = selectionMgr.GetSelectedObject6(1, -1) as SketchPoint;
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select right bottom slot corner");
            SketchPoint rightBottomSlotCorner = selectionMgr.GetSelectedObject6(1, -1) as SketchPoint;
            swModel.ClearSelection2(true);

            if (leftBottomSlotCorner == null || rightBottomSlotCorner == null)
                throw new Exception("Could not access distance-sheet slot corner points");

            SketchSegment line3 = (SketchSegment)swSketchManager.CreateLine(-halfWidth, bottomY, 0, 0, 0, 0);
            SketchSegment line4 = (SketchSegment)swSketchManager.CreateLine(halfWidth, bottomY, 0, 0, 0, 0);
            if (line3 == null || line4 == null)
                throw new Exception("Could not create distance-sheet corner lines");

            SketchLine line3Sketch = line3 as SketchLine;
            SketchLine line4Sketch = line4 as SketchLine;
            if (line3Sketch == null || line4Sketch == null)
                throw new Exception("Could not access distance-sheet corner line geometry");

            SketchPoint line3Start = line3Sketch.GetStartPoint2();
            SketchPoint line3End = line3Sketch.GetEndPoint2();
            SketchPoint line4Start = line4Sketch.GetStartPoint2();
            SketchPoint line4End = line4Sketch.GetEndPoint2();
            if (line3Start == null || line3End == null || line4Start == null || line4End == null)
                throw new Exception("Could not access distance-sheet corner line endpoints");

            SketchPoint line3OriginPoint = Math.Abs(line3Start.X) < 0.000001 && Math.Abs(line3Start.Y) < 0.000001 ? line3Start : line3End;
            SketchPoint line4OriginPoint = Math.Abs(line4Start.X) < 0.000001 && Math.Abs(line4Start.Y) < 0.000001 ? line4Start : line4End;
            SketchPoint line3CornerPoint = line3OriginPoint == line3Start ? line3End : line3Start;
            SketchPoint line4CornerPoint = line4OriginPoint == line4Start ? line4End : line4Start;

            if (!line3OriginPoint.Select4(false, null))
                throw new Exception("Could not select left slot line origin point");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for left slot line");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!line4OriginPoint.Select4(false, null))
                throw new Exception("Could not select right slot line origin point");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for right slot line");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!line3CornerPoint.Select4(false, null))
                throw new Exception("Could not select left slot corner point");
            if (!leftBottomSlotCorner.Select4(true, null))
                throw new Exception("Could not select left bottom slot corner");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!line4CornerPoint.Select4(false, null))
                throw new Exception("Could not select right slot corner point");
            if (!rightBottomSlotCorner.Select4(true, null))
                throw new Exception("Could not select right bottom slot corner");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            // Turn the slot sketch into a real cut feature.
            swSketchManager.InsertSketch(true);
            if (!SelectSketchByIndex(swModel, 2))
                throw new Exception("Could not select distance-sheet slot sketch");
            Feature slotCutFeature = swModel.FeatureManager.FeatureCut4(
                false, false, false,
                (int)swEndConditions_e.swEndCondThroughAll, (int)swEndConditions_e.swEndCondThroughAll,
                0, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                false, true, true, true, true, false,
                0, 0, false, false);
            if (slotCutFeature == null)
                throw new Exception("Failed to create distance-sheet slot cut extrusion");

            // Repeat the one slot feature around the center axis.
            swModel.Extension.SelectByID2("Z-Achse", "AXIS", 0, 0, 0, false, 1, null, 0);
            swModel.Extension.SelectByID2("Cut-Extrude1", "BODYFEATURE", 0, 0, 0, true, 4, null, 0);
            Feature myPattern = (Feature)swModel.FeatureManager.FeatureCircularPattern5(
                slotPatternCount,
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
                false
            );
            if (myPattern == null)
                throw new Exception("Failed to create distance-sheet circular pattern");

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "FACE", (outerRadius + innerRadius) / 2.0, 0, plateThickness, false, 0, null, 0);
            if (!selected) throw new Exception("Could not reselect top face for boss sketch");

            bool bossSketchInferenceWasEnabled = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            swSketchManager.InsertSketch(true);

            // Phase 3:
            // Build one rectangular boss using a rotated local coordinate system.
            double bossTopCenterX = bossDirectionX * bossTopCenterRadius;
            double bossTopCenterY = bossDirectionY * bossTopCenterRadius;
            double bossBottomCenterX = bossDirectionX * bossBottomCenterRadius;
            double bossBottomCenterY = bossDirectionY * bossBottomCenterRadius;

            double topLeftX = bossTopCenterX - bossNormalX * bossHalfWidth;
            double topLeftY = bossTopCenterY - bossNormalY * bossHalfWidth;
            double topRightX = bossTopCenterX + bossNormalX * bossHalfWidth;
            double topRightY = bossTopCenterY + bossNormalY * bossHalfWidth;
            double bottomRightX = bossBottomCenterX + bossNormalX * bossHalfWidth;
            double bottomRightY = bossBottomCenterY + bossNormalY * bossHalfWidth;
            double bottomLeftX = bossBottomCenterX - bossNormalX * bossHalfWidth;
            double bottomLeftY = bossBottomCenterY - bossNormalY * bossHalfWidth;

            SketchSegment bossTopEdge = (SketchSegment)swSketchManager.CreateLine(topLeftX, topLeftY, 0, topRightX, topRightY, 0);
            SketchSegment bossRightEdge = (SketchSegment)swSketchManager.CreateLine(topRightX, topRightY, 0, bottomRightX, bottomRightY, 0);
            SketchSegment bossBottomEdge = (SketchSegment)swSketchManager.CreateLine(bottomRightX, bottomRightY, 0, bottomLeftX, bottomLeftY, 0);
            SketchSegment bossLeftEdge = (SketchSegment)swSketchManager.CreateLine(bottomLeftX, bottomLeftY, 0, topLeftX, topLeftY, 0);
            if (bossTopEdge == null || bossRightEdge == null || bossBottomEdge == null || bossLeftEdge == null)
                throw new Exception("Could not create boss rectangle");
            SketchSegment bossConstructionLine = (SketchSegment)swSketchManager.CreateCenterLine(0, 0, 0, bossTopCenterX, bossTopCenterY, 0);
            if (bossConstructionLine == null)
                throw new Exception("Could not create boss construction line");
            SketchLine bossConstructionSketchLine = bossConstructionLine as SketchLine;
            if (bossConstructionSketchLine == null)
                throw new Exception("Could not access boss construction line geometry");
            SketchPoint bossConstructionStartPoint = bossConstructionSketchLine.GetStartPoint2();
            SketchPoint bossConstructionEndPoint = bossConstructionSketchLine.GetEndPoint2();
            if (bossConstructionStartPoint == null || bossConstructionEndPoint == null)
                throw new Exception("Could not access boss construction line points");

            swModel.ClearSelection2(true);
            if (!bossConstructionStartPoint.Select4(false, null))
                throw new Exception("Could not select boss construction line start point");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for boss construction line");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!bossConstructionEndPoint.Select4(false, null))
                throw new Exception("Could not select boss construction line end point");
            if (!bossTopEdge.Select4(true, null))
                throw new Exception("Could not select boss top edge for construction-line midpoint relation");
            swModel.SketchAddConstraints("sgATMIDDLE");
            swModel.ClearSelection2(true);

            if (!bossConstructionLine.Select4(false, null))
                throw new Exception("Could not select boss construction line for parallel relation");
            if (!bossRightEdge.Select4(true, null))
                throw new Exception("Could not select boss right edge for construction-line parallel relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            if (!bossTopEdge.Select4(false, null))
                throw new Exception("Could not select boss top edge for perpendicular relation");
            if (!bossRightEdge.Select4(true, null))
                throw new Exception("Could not select boss right edge for perpendicular relation");
            swModel.SketchAddConstraints("sgPERPENDICULAR");
            swModel.ClearSelection2(true);

            if (!bossRightEdge.Select4(false, null))
                throw new Exception("Could not select boss right edge for left-edge relation");
            if (!bossLeftEdge.Select4(true, null))
                throw new Exception("Could not select boss left edge for right-edge relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            if (!bossBottomEdge.Select4(false, null))
                throw new Exception("Could not select boss bottom edge for parallel relation");
            if (!bossTopEdge.Select4(true, null))
                throw new Exception("Could not select boss top edge for bottom-edge relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            // These dimensions fully define the boss size, angle, and radial location.
            if (!bossRightEdge.Select4(false, null))
                throw new Exception("Could not select boss right edge for angle dimension");
            selected = swModel.Extension.SelectByID2("Y-Achse", "AXIS", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select Y axis for boss angle dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(bossTopCenterX + 0.04, bossTopCenterY - 0.04, 0);
            if (displayDim == null) throw new Exception("Could not create boss rectangle angle dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss rectangle angle dimension");
            swDim.SystemValue = bossCenterlineAngleRadians;
            swModel.ClearSelection2(true);

            if (!bossTopEdge.Select4(false, null))
                throw new Exception("Could not select boss top edge for width dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(
                bossTopCenterX + bossNormalX * 0.025,
                bossTopCenterY + bossNormalY * 0.025,
                0);
            if (displayDim == null) throw new Exception("Could not create boss width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss width dimension");
            swDim.SystemValue = bossRectangleWidth;
            swModel.ClearSelection2(true);

            if (!bossRightEdge.Select4(false, null))
                throw new Exception("Could not select boss right edge for height dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(
                ((topRightX + bottomRightX) / 2.0) + bossNormalX * 0.025,
                ((topRightY + bottomRightY) / 2.0) + bossNormalY * 0.025,
                0);
            if (displayDim == null) throw new Exception("Could not create boss height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss height dimension");
            swDim.SystemValue = bossRectangleHeight;
            swModel.ClearSelection2(true);

            if (!bossTopEdge.Select4(false, null))
                throw new Exception("Could not select boss top edge for top offset dimension");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for top offset dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(
                bossTopCenterX / 2.0 + bossNormalX * 0.02,
                bossTopCenterY / 2.0 + bossNormalY * 0.02,
                0);
            if (displayDim == null) throw new Exception("Could not create boss top offset dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss top offset dimension");
            // Position the top edge at outer-radius minus the requested clearance.
            swDim.SystemValue = bossTopCenterRadius;
            swModel.ClearSelection2(true);

            swSketchManager.InsertSketch(true);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, bossSketchInferenceWasEnabled);
            if (!SelectSketchByIndex(swModel, 3))
                throw new Exception("Could not select boss sketch");
            Feature bossFeature = swModel.FeatureManager.FeatureExtrusion2(
                true, false, false,
                (int)swEndConditions_e.swEndCondBlind, 0,
                bossExtrusionDepth, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                false, true, true, 0, 0, false);
            if (bossFeature == null)
                throw new Exception("Failed to create separate boss body");

            // Phase 4:
            // Add the relief cut on the boss front face.
            object[] bossFeatureFaces = bossFeature.GetFaces() as object[];
            Face2 bossFrontFace = null;
            double bossFaceTolerance = 0.00001;
            double expectedBossFrontFaceArea = bossRectangleWidth * bossExtrusionDepth;
            double expectedBossFrontFaceCenterX = bossTopCenterX;
            double expectedBossFrontFaceCenterY = bossTopCenterY;
            double bossFrontFaceAreaTolerance = expectedBossFrontFaceArea * 0.1;
            double bestBossFrontFaceCenterDelta = double.MaxValue;
            if (bossFeatureFaces != null)
            {
                foreach (object faceObj in bossFeatureFaces)
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

                    if (Math.Abs(candidateBox[2] - plateThickness) < bossFaceTolerance
                        && Math.Abs(candidateBox[5] - (plateThickness + bossExtrusionDepth)) < bossFaceTolerance)
                    {
                        double candidateAreaDelta = Math.Abs(candidateFace.GetArea() - expectedBossFrontFaceArea);
                        if (candidateAreaDelta > bossFrontFaceAreaTolerance)
                            continue;

                        double candidateCenterX = (candidateBox[0] + candidateBox[3]) / 2.0;
                        double candidateCenterY = (candidateBox[1] + candidateBox[4]) / 2.0;
                        double candidateCenterDelta = Math.Sqrt(
                            Math.Pow(candidateCenterX - expectedBossFrontFaceCenterX, 2)
                            + Math.Pow(candidateCenterY - expectedBossFrontFaceCenterY, 2));
                        if (candidateCenterDelta < bestBossFrontFaceCenterDelta)
                        {
                            bestBossFrontFaceCenterDelta = candidateCenterDelta;
                            bossFrontFace = candidateFace;
                        }
                    }
                }
            }

            if (bossFrontFace == null)
                throw new Exception("Could not find the front face of the rectangular boss for the cut sketch");

            swModel.ClearSelection2(true);
            Entity bossFrontFaceEntity = bossFrontFace as Entity;
            if (bossFrontFaceEntity == null || !bossFrontFaceEntity.Select4(false, null))
                throw new Exception("Could not select the front face of the rectangular boss for the cut sketch");
            bool bossCutSketchInferenceWasEnabled = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            swSketchManager.InsertSketch(true);

            if (bossCutOuterTabWidth <= 0
                || bossCutOuterTabHeight <= 0
                || bossCutTopShelfThickness <= 0
                || bossCutInnerLegWidth <= 0
                || bossCutBoundaryExtension <= 0)
                throw new Exception("Boss cut dimensions must be positive");

            if (bossCutOuterTabHeight <= bossCutTopShelfThickness)
                throw new Exception("Boss cut outer tab height must exceed the top shelf thickness");

            if (bossCutOuterTabHeight >= bossRectangleWidth
                || bossCutTopShelfThickness >= bossRectangleWidth)
                throw new Exception("Boss cut vertical dimensions must fit within the boss width");

            if (bossCutOuterTabWidth >= bossExtrusionDepth
                || bossCutInnerLegWidth >= bossExtrusionDepth
                || bossCutOuterTabWidth >= (bossExtrusionDepth - bossCutInnerLegWidth))
                throw new Exception("Boss cut horizontal dimensions must fit within the boss extrusion depth");

            MathUtility mathUtility = (MathUtility)_swApp.GetMathUtility();
            if (mathUtility == null)
                throw new Exception("Could not access SolidWorks math utility for boss cut sketch");
            Sketch activeBossCutSketch = swModel.GetActiveSketch2() as Sketch;
            if (activeBossCutSketch == null)
                throw new Exception("Could not access active boss cut sketch");
            MathTransform modelToBossCutSketchTransform = activeBossCutSketch.ModelToSketchTransform as MathTransform;
            if (modelToBossCutSketchTransform == null)
                throw new Exception("Could not access boss cut sketch transform");

            void CreateBossFrontFaceModelPoint(double xLocal, double yLocal, out double x, out double y, out double z)
            {
                x = topLeftX + bossNormalX * yLocal;
                y = topLeftY + bossNormalY * yLocal;
                z = plateThickness + bossExtrusionDepth - xLocal;
            }

            double[] ConvertBossCutPointToSketchCoordinates(double modelX, double modelY, double modelZ)
            {
                MathPoint modelPoint = (MathPoint)mathUtility.CreatePoint(new double[] { modelX, modelY, modelZ });
                if (modelPoint == null)
                    throw new Exception("Could not create boss cut model point");
                MathPoint sketchPoint = (MathPoint)modelPoint.MultiplyTransform(modelToBossCutSketchTransform);
                if (sketchPoint == null)
                    throw new Exception("Could not transform boss cut point into sketch coordinates");
                double[] sketchPointData = sketchPoint.ArrayData as double[];
                if (sketchPointData == null || sketchPointData.Length < 3)
                    throw new Exception("Could not read transformed boss cut sketch coordinates");
                return sketchPointData;
            }

            double[] GetSketchSegmentMidpoint(SketchSegment sketchSegment)
            {
                SketchLine sketchLine = sketchSegment as SketchLine;
                if (sketchLine == null)
                    throw new Exception("Boss cut reference segment is not a sketch line");
                SketchPoint startPoint = sketchLine.GetStartPoint2();
                SketchPoint endPoint = sketchLine.GetEndPoint2();
                if (startPoint == null || endPoint == null)
                    throw new Exception("Could not access boss cut reference segment endpoints");
                return new double[]
                {
                    (startPoint.X + endPoint.X) / 2.0,
                    (startPoint.Y + endPoint.Y) / 2.0
                };
            }

            SketchSegment FindClosestBossCutReferenceSegment(object[] referenceSegments, double[] targetMidpoint, string label)
            {
                SketchSegment bestSegment = null;
                double bestDistance = double.MaxValue;
                foreach (object segmentObj in referenceSegments)
                {
                    SketchSegment candidateSegment = segmentObj as SketchSegment;
                    if (candidateSegment == null)
                        continue;
                    double[] candidateMidpoint = GetSketchSegmentMidpoint(candidateSegment);
                    double candidateDistance = Math.Sqrt(
                        Math.Pow(candidateMidpoint[0] - targetMidpoint[0], 2)
                        + Math.Pow(candidateMidpoint[1] - targetMidpoint[1], 2));
                    if (candidateDistance < bestDistance)
                    {
                        bestDistance = candidateDistance;
                        bestSegment = candidateSegment;
                    }
                }

                if (bestSegment == null)
                    throw new Exception($"Could not find boss cut reference segment for {label}");

                return bestSegment;
            }

            double bossCutProfileWidth = bossExtrusionDepth;
            double bossCutProfileHeight = bossRectangleWidth;
            double bossCutInnerLegStartX = bossCutProfileWidth - bossCutInnerLegWidth;
            double bossCutOuterX = -bossCutBoundaryExtension;
            double bossCutMirroredExtensionY = -bossCutBoundaryExtension;
            double bossCutMirroredShelfY = bossCutProfileHeight - bossCutTopShelfThickness;
            double bossCutMirroredTabY = bossCutProfileHeight - bossCutOuterTabHeight;

            swModel.ClearSelection2(true);
            if (!bossFrontFaceEntity.Select4(false, null))
                throw new Exception("Could not reselect boss front face for cut-sketch references");
            if (!swSketchManager.SketchUseEdge3(false, false))
                throw new Exception("Could not convert boss front face edges into cut-sketch references");
            object[] bossCutReferenceSegments = activeBossCutSketch.GetSketchSegments() as object[];
            if (bossCutReferenceSegments == null || bossCutReferenceSegments.Length < 4)
                throw new Exception("Could not access boss front face reference geometry in the cut sketch");
            foreach (object segmentObj in bossCutReferenceSegments)
            {
                SketchSegment referenceSegment = segmentObj as SketchSegment;
                if (referenceSegment == null)
                    continue;
                referenceSegment.ConstructionGeometry = true;
            }
            swModel.ClearSelection2(true);

            CreateBossFrontFaceModelPoint(0, bossCutProfileHeight / 2.0, out double outerBoundaryMidModelX, out double outerBoundaryMidModelY, out double outerBoundaryMidModelZ);
            CreateBossFrontFaceModelPoint(bossCutProfileWidth, bossCutProfileHeight / 2.0, out double innerBoundaryMidModelX, out double innerBoundaryMidModelY, out double innerBoundaryMidModelZ);
            CreateBossFrontFaceModelPoint(bossCutProfileWidth / 2.0, 0, out double baseBoundaryMidModelX, out double baseBoundaryMidModelY, out double baseBoundaryMidModelZ);
            CreateBossFrontFaceModelPoint(bossCutProfileWidth / 2.0, bossCutProfileHeight, out double topBoundaryMidModelX, out double topBoundaryMidModelY, out double topBoundaryMidModelZ);

            double[] outerBoundaryMid = ConvertBossCutPointToSketchCoordinates(outerBoundaryMidModelX, outerBoundaryMidModelY, outerBoundaryMidModelZ);
            double[] innerBoundaryMid = ConvertBossCutPointToSketchCoordinates(innerBoundaryMidModelX, innerBoundaryMidModelY, innerBoundaryMidModelZ);
            double[] baseBoundaryMid = ConvertBossCutPointToSketchCoordinates(baseBoundaryMidModelX, baseBoundaryMidModelY, baseBoundaryMidModelZ);
            double[] topBoundaryMid = ConvertBossCutPointToSketchCoordinates(topBoundaryMidModelX, topBoundaryMidModelY, topBoundaryMidModelZ);
            SketchSegment outerBoundarySegment = FindClosestBossCutReferenceSegment(bossCutReferenceSegments, outerBoundaryMid, "outer boundary");
            SketchSegment innerBoundarySegment = FindClosestBossCutReferenceSegment(bossCutReferenceSegments, innerBoundaryMid, "inner boundary");
            SketchSegment baseBoundarySegment = FindClosestBossCutReferenceSegment(bossCutReferenceSegments, baseBoundaryMid, "base boundary");
            SketchSegment topBoundarySegment = FindClosestBossCutReferenceSegment(bossCutReferenceSegments, topBoundaryMid, "top boundary");

            CreateBossFrontFaceModelPoint(bossCutOuterX, bossCutMirroredExtensionY, out double cutP1ModelX, out double cutP1ModelY, out double cutP1ModelZ);
            CreateBossFrontFaceModelPoint(bossCutInnerLegStartX, bossCutMirroredExtensionY, out double cutP2ModelX, out double cutP2ModelY, out double cutP2ModelZ);
            CreateBossFrontFaceModelPoint(bossCutInnerLegStartX, bossCutMirroredShelfY, out double cutP3ModelX, out double cutP3ModelY, out double cutP3ModelZ);
            CreateBossFrontFaceModelPoint(bossCutOuterTabWidth, bossCutMirroredShelfY, out double cutP4ModelX, out double cutP4ModelY, out double cutP4ModelZ);
            CreateBossFrontFaceModelPoint(bossCutOuterTabWidth, bossCutMirroredTabY, out double cutP5ModelX, out double cutP5ModelY, out double cutP5ModelZ);
            CreateBossFrontFaceModelPoint(bossCutOuterX, bossCutMirroredTabY, out double cutP6ModelX, out double cutP6ModelY, out double cutP6ModelZ);

            double[] cutP1 = ConvertBossCutPointToSketchCoordinates(cutP1ModelX, cutP1ModelY, cutP1ModelZ);
            double[] cutP2 = ConvertBossCutPointToSketchCoordinates(cutP2ModelX, cutP2ModelY, cutP2ModelZ);
            double[] cutP3 = ConvertBossCutPointToSketchCoordinates(cutP3ModelX, cutP3ModelY, cutP3ModelZ);
            double[] cutP4 = ConvertBossCutPointToSketchCoordinates(cutP4ModelX, cutP4ModelY, cutP4ModelZ);
            double[] cutP5 = ConvertBossCutPointToSketchCoordinates(cutP5ModelX, cutP5ModelY, cutP5ModelZ);
            double[] cutP6 = ConvertBossCutPointToSketchCoordinates(cutP6ModelX, cutP6ModelY, cutP6ModelZ);

            SketchSegment bossCutLine1 = (SketchSegment)swSketchManager.CreateLine(cutP1[0], cutP1[1], 0, cutP2[0], cutP2[1], 0);
            SketchSegment bossCutLine2 = (SketchSegment)swSketchManager.CreateLine(cutP2[0], cutP2[1], 0, cutP3[0], cutP3[1], 0);
            SketchSegment bossCutLine3 = (SketchSegment)swSketchManager.CreateLine(cutP3[0], cutP3[1], 0, cutP4[0], cutP4[1], 0);
            SketchSegment bossCutLine4 = (SketchSegment)swSketchManager.CreateLine(cutP4[0], cutP4[1], 0, cutP5[0], cutP5[1], 0);
            SketchSegment bossCutLine5 = (SketchSegment)swSketchManager.CreateLine(cutP5[0], cutP5[1], 0, cutP6[0], cutP6[1], 0);
            SketchSegment bossCutLine6 = (SketchSegment)swSketchManager.CreateLine(cutP6[0], cutP6[1], 0, cutP1[0], cutP1[1], 0);
            if (bossCutLine1 == null
                || bossCutLine2 == null
                || bossCutLine3 == null
                || bossCutLine4 == null
                || bossCutLine5 == null
                || bossCutLine6 == null)
                throw new Exception("Could not create boss cut profile");

            swModel.ClearSelection2(true);
            if (!bossCutLine1.Select4(false, null))
                throw new Exception("Could not select boss cut top line for parallel relation");
            if (!baseBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut base reference for mirrored parallel relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            if (!bossCutLine3.Select4(false, null))
                throw new Exception("Could not select boss cut shelf line for parallel relation");
            if (!topBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut top reference for mirrored shelf relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            if (!bossCutLine5.Select4(false, null))
                throw new Exception("Could not select boss cut tab line for parallel relation");
            if (!topBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut top reference for mirrored tab relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            if (!bossCutLine2.Select4(false, null))
                throw new Exception("Could not select boss cut inner leg for parallel relation");
            if (!outerBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut outer-face reference for inner-leg relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            if (!bossCutLine4.Select4(false, null))
                throw new Exception("Could not select boss cut outer tab leg for parallel relation");
            if (!outerBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut outer-face reference for tab-leg relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            if (!bossCutLine6.Select4(false, null))
                throw new Exception("Could not select boss cut extension leg for parallel relation");
            if (!outerBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut outer-face reference for extension-leg relation");
            swModel.SketchAddConstraints("sgPARALLEL");
            swModel.ClearSelection2(true);

            if (!bossCutLine4.Select4(false, null))
                throw new Exception("Could not select boss cut outer tab leg for width dimension");
            if (!outerBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut outer-face reference for width dimension");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(
                (cutP4[0] + outerBoundaryMid[0]) / 2.0,
                ((cutP4[1] + cutP5[1]) / 2.0) - 0.01,
                0);
            if (displayDim == null) throw new Exception("Could not create boss cut outer tab width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss cut outer tab width dimension");
            swDim.SystemValue = bossCutOuterTabWidth;
            swModel.ClearSelection2(true);

            if (!bossCutLine2.Select4(false, null))
                throw new Exception("Could not select boss cut inner leg for width dimension");
            if (!innerBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut inner-face reference for width dimension");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(
                (cutP2[0] + innerBoundaryMid[0]) / 2.0,
                ((cutP2[1] + cutP3[1]) / 2.0) - 0.01,
                0);
            if (displayDim == null) throw new Exception("Could not create boss cut inner leg width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss cut inner leg width dimension");
            swDim.SystemValue = bossCutInnerLegWidth;
            swModel.ClearSelection2(true);

            if (!bossCutLine3.Select4(false, null))
                throw new Exception("Could not select boss cut shelf line for height dimension");
            if (!topBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut top reference for mirrored shelf height dimension");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(
                ((cutP3[0] + cutP4[0]) / 2.0) + 0.01,
                (cutP3[1] + topBoundaryMid[1]) / 2.0,
                0);
            if (displayDim == null) throw new Exception("Could not create boss cut shelf height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss cut shelf height dimension");
            swDim.SystemValue = bossCutTopShelfThickness;
            swModel.ClearSelection2(true);

            if (!bossCutLine5.Select4(false, null))
                throw new Exception("Could not select boss cut tab line for height dimension");
            if (!topBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut top reference for mirrored tab height dimension");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(
                ((cutP5[0] + cutP6[0]) / 2.0) + 0.01,
                (cutP5[1] + topBoundaryMid[1]) / 2.0,
                0);
            if (displayDim == null) throw new Exception("Could not create boss cut tab height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss cut tab height dimension");
            swDim.SystemValue = bossCutOuterTabHeight;
            swModel.ClearSelection2(true);

            if (!bossCutLine6.Select4(false, null))
                throw new Exception("Could not select boss cut outer extension line for boundary dimension");
            if (!outerBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut outer-face reference for boundary extension dimension");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(
                (cutP6[0] + outerBoundaryMid[0]) / 2.0,
                ((cutP6[1] + cutP1[1]) / 2.0) + 0.01,
                0);
            if (displayDim == null) throw new Exception("Could not create boss cut outer boundary extension dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss cut outer boundary extension dimension");
            swDim.SystemValue = bossCutBoundaryExtension;
            swModel.ClearSelection2(true);

            if (!bossCutLine1.Select4(false, null))
                throw new Exception("Could not select boss cut top extension line for boundary dimension");
            if (!baseBoundarySegment.Select4(true, null))
                throw new Exception("Could not select boss cut base-face reference for mirrored boundary extension dimension");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(
                ((cutP1[0] + cutP2[0]) / 2.0) - 0.01,
                (cutP1[1] + baseBoundaryMid[1]) / 2.0,
                0);
            if (displayDim == null) throw new Exception("Could not create boss cut top boundary extension dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss cut top boundary extension dimension");
            swDim.SystemValue = bossCutBoundaryExtension;
            swModel.ClearSelection2(true);

            swSketchManager.InsertSketch(true);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, bossCutSketchInferenceWasEnabled);
            Feature CreateBossProfileCut(bool reverseDirection)
            {
                swModel.ClearSelection2(true);
                if (!SelectSketchByIndex(swModel, 4))
                    throw new Exception("Could not select boss cut sketch");
                return swModel.FeatureManager.FeatureCut4(
                    true, false, reverseDirection,
                    (int)swEndConditions_e.swEndCondThroughAll, 0,
                    0, 0,
                    false, false, false, false,
                    0, 0, false, false, false, false,
                    false, true, true, true, true, false,
                    0, 0, false, false);
            }

            Feature bossCutFeature = CreateBossProfileCut(false);
            if (bossCutFeature == null)
                bossCutFeature = CreateBossProfileCut(true);
            if (bossCutFeature == null)
                throw new Exception("Failed to create rectangular-boss profile cut");

            if (bossCircularPatternCount < 2)
                throw new Exception("Boss circular pattern count must be at least 2");

            PartDoc distanceSheetPart = swModel as PartDoc;
            if (distanceSheetPart == null)
                throw new Exception("Could not access part document for boss body pattern");

            object[] solidBodies = distanceSheetPart.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
            Body2 patternedBossBody = null;
            double expectedBossBodyCenterX = (bossTopCenterX + bossBottomCenterX) / 2.0;
            double expectedBossBodyCenterY = (bossTopCenterY + bossBottomCenterY) / 2.0;
            double bestBossBodyCenterDelta = double.MaxValue;
            if (solidBodies != null)
            {
                foreach (object bodyObj in solidBodies)
                {
                    Body2 candidateBody = bodyObj as Body2;
                    if (candidateBody == null)
                        continue;

                    double[] candidateBodyBox = candidateBody.GetBodyBox() as double[];
                    if (candidateBodyBox == null || candidateBodyBox.Length < 6)
                        continue;

                    if (Math.Abs(candidateBodyBox[2] - plateThickness) >= bossFaceTolerance
                        || Math.Abs(candidateBodyBox[5] - (plateThickness + bossExtrusionDepth)) >= bossFaceTolerance)
                        continue;

                    double candidateBodyCenterX = (candidateBodyBox[0] + candidateBodyBox[3]) / 2.0;
                    double candidateBodyCenterY = (candidateBodyBox[1] + candidateBodyBox[4]) / 2.0;
                    double candidateCenterDelta = Math.Sqrt(
                        Math.Pow(candidateBodyCenterX - expectedBossBodyCenterX, 2)
                        + Math.Pow(candidateBodyCenterY - expectedBossBodyCenterY, 2));
                    if (candidateCenterDelta < bestBossBodyCenterDelta)
                    {
                        bestBossBodyCenterDelta = candidateCenterDelta;
                        patternedBossBody = candidateBody;
                    }
                }
            }

            if (patternedBossBody == null)
                throw new Exception("Could not find the finished rectangular boss body for circular pattern");

            selectionMgr = swModel.SelectionManager as SelectionMgr;
            if (selectionMgr == null)
                throw new Exception("Could not access selection manager for boss body pattern");
            SelectData bodyPatternSelectData = selectionMgr.CreateSelectData();
            if (bodyPatternSelectData == null)
                throw new Exception("Could not create selection data for boss body pattern");
            bodyPatternSelectData.Mark = 256;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("Z-Achse", "AXIS", 0, 0, 0, false, 1, null, 0);
            if (!selected) throw new Exception("Could not select Z axis for boss body circular pattern");
            if (!patternedBossBody.Select2(true, bodyPatternSelectData))
                throw new Exception("Could not select finished boss body for circular pattern");
            Feature bossBodyPattern = (Feature)swModel.FeatureManager.FeatureCircularPattern5(
                bossCircularPatternCount,
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
                false
            );
            if (bossBodyPattern == null)
                throw new Exception("Failed to create circular pattern of the rectangular boss body");

            swModel.ShowNamedView2("*Front", (int)swStandardViews_e.swFrontView);
            swModel.ViewZoomtofit2();

            // Save the completed part using the chosen local-or-PDM workflow.
            string savedPath;
            if (saveToPdm)
            {
                savedPath = _pdm.SaveAsPdm(swModel);
                Console.WriteLine($"Part saved to PDM: {savedPath}");
            }
            else
            {
                savedPath = Path.Combine(outFolder, GetRequiredLocalFileName());
                swModel.SaveAs3(savedPath, 0, 1);
                Console.WriteLine($"Part saved locally: {savedPath}");
            }

            if (shouldCloseAfterCreate)
            {
                // Use GetTitle() to ensure we close the specific document we just saved
                _swApp.CloseDoc(swModel.GetTitle());
                Console.WriteLine("Part closed after creating.");
            }

            if (saveToPdm)
            {
                _pdm.UpdateBirrDataCard(savedPath, PdmDataCard.ToDictionary());
            }

            Console.WriteLine("Done!");

            return Path.GetFileName(savedPath); 
        }
    }

}





