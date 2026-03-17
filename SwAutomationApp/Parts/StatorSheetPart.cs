using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Creates the main stator sheet part.
///
/// Geometrically, this part is:
/// - an annular base ring
/// - one slot cut
/// - one guide/boss detail around the slot
/// - a circular pattern that repeats the slot geometry around the part
/// </summary>
public sealed class StatorSheetPart
{
    
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    public StatorSheetPart(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        _pdm = pdm ?? throw new ArgumentNullException(nameof(pdm));
    }

    // File and save settings.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "StatorBleche.SLDPRT";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    // Main editable geometry values.
    public double OuterDiameter { get; set; } = 0.99;
    public double InnerDiameter { get; set; } = 0.64;
    public double PlateThickness { get; set; } = 0.1;
    public double SlotWidth { get; set; } = 0.0157;
    public double SlotBottomY { get; set; } = 0.32;
    public double SlotTopY { get; set; } = 0.4052;
    public double AngleInDegrees { get; set; } = 70.0;
    public double FilletRadius { get; set; } = 0.001;
    public double SlotGuideSpacing { get; set; } = 0.0057;
    public double SlotGuideOffset { get; set; } = -0.0007;
    public int SlotPatternCount { get; set; } = 60;
    public string MaterialName { get; set; } = "AISI 1020";

    private string GetRequiredOutputFolder() => AutomationSupport.RequireText(OutputFolder, nameof(OutputFolder), nameof(StatorSheetPart));
    private string GetRequiredLocalFileName() => AutomationSupport.RequireText(LocalFileName, nameof(LocalFileName), nameof(StatorSheetPart));
    private AutomationUiScope BeginAutomationUiSuppression() => new(_swApp);

    /// <summary>
    /// Creates the stator sheet model and saves it.
    /// </summary>
    public string Create()
    {
        
        using var automationUi = BeginAutomationUiSuppression();

        // Read the current object settings first so the method works from one clear snapshot.
        string outFolder = GetRequiredOutputFolder();
        bool closeAfterCreate = CloseAfterCreate;
        bool saveToPdm = SaveToPdm;

        // Main dimensions (m) - change these only.
        double outerDiameter = OuterDiameter;
        double innerDiameter = InnerDiameter;
        double plateThicknessValue = PlateThickness;
        double slotWidth = SlotWidth;
        double slotBottomY = SlotBottomY;
        double slotTopY = SlotTopY;
        double angleInDegrees = AngleInDegrees;
        double filletradius = FilletRadius;
        double slotGuideSpacingValue = SlotGuideSpacing;
        double slotGuideOffsetValue = SlotGuideOffset;
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
        double angleInRadians = angleInDegrees * Math.PI / 180.0;
        double filletRadius = filletradius;
        double slotGuideSpacing = slotGuideSpacingValue;
        double slotGuideOffset = slotGuideOffsetValue;

        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        try
        {
            Dimension swDim = null;
            DisplayDimension displayDim = null;

            if (!Directory.Exists(outFolder))
                Directory.CreateDirectory(outFolder);

            // Create a new blank part from the user's default part template.
            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            if (swModel == null)
                throw new Exception("Failed to create new part");

            swSketchManager = swModel.SketchManager;
            SelectionMgr selMgr = swModel.SelectionManager;

            // Apply the chosen material name to the generated part.
            PartDoc statorPart = swModel as PartDoc;
            statorPart.SetMaterialPropertyName2("", "", Name: materialName);

            // Phase 1:
            // Create the base annulus sketch and extrude it into the main plate thickness.
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
            swModel.Extension.SelectByID2("Skizze1", "SKETCH", 0.4, 0, 0, false, 0, null, 0);
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

            // Phase 2:
            // Sketch the main slot rectangle on the top face and dimension its width, height,
            // and radial position.
            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCenterRectangle(0, centerY, 0, halfWidth, topY, 0);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom edge");
            displayDim = (DisplayDimension)swModel.AddDimension2(leftX + 0.01, bottomY - 0.01, 0);
            if (displayDim == null) throw new Exception("Could not create slot width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot width dimension");
            swDim.SystemValue = halfWidth * 2;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom-right point");
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, topY, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot top-right point");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(halfWidth + 0.05, centerY, 0);
            if (displayDim == null) throw new Exception("Could not create slot height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot height dimension");
            swDim.SystemValue = topY - bottomY;

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", 0, centerY, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgVERTICALPOINTS2D");
            swModel.ClearSelection2(true);

            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY / 2, 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = innerRadius;
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);

            // Phase 3:
            // Build the slot guide geometry using helper lines and a fillet.
            SketchSegment line1 = (SketchSegment)swSketchManager.CreateLine(0.02, 0.02, 0, 0.05, 0.05, 0);
            SketchSegment line2 = (SketchSegment)swSketchManager.CreateLine(0.05, 0.05, 0, 0.02, 0.05, 0);
            if (line1 == null || line2 == null) throw new Exception("Could not create slot guide lines");
            SketchLine line1Sketch = line1 as SketchLine;
            SketchLine line2Sketch = line2 as SketchLine;
            if (line1Sketch == null || line2Sketch == null) throw new Exception("Could not access slot guide line geometry");
            SketchPoint line1StartPoint = line1Sketch.GetStartPoint2();
            SketchPoint line1EndPoint = line1Sketch.GetEndPoint2();
            SketchPoint line2StartPoint = line2Sketch.GetStartPoint2();
            SketchPoint line2EndPoint = line2Sketch.GetEndPoint2();
            if (line1StartPoint == null || line1EndPoint == null || line2StartPoint == null || line2EndPoint == null)
                throw new Exception("Could not access slot guide endpoints");
            SketchPoint lowerWallPoint = line1StartPoint.Y <= line1EndPoint.Y ? line1StartPoint : line1EndPoint;
            SketchPoint upperWallPoint = line2StartPoint.X <= line2EndPoint.X ? line2StartPoint : line2EndPoint;

            swModel.Extension.SelectByID2("", "SKETCHPOINT", 0.02, 0.02, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, bottomY + 0.05, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");

            swModel.Extension.SelectByID2("", "SKETCHPOINT", 0.02, 0.05, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, bottomY + 0.05, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!line2.Select4(false, null))
                throw new Exception("Could not select upper slot guide segment for fillet");
            if (!line1.Select4(true, null))
                throw new Exception("Could not select lower slot guide segment for fillet");
            SketchSegment slotfilletArc = swSketchManager.CreateFillet(filletRadius, (int)swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry);
            if (slotfilletArc == null) throw new Exception("Could not create slot fillet");

            line2.Select4(false, null);
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            if (!upperWallPoint.Select4(false, null))
                throw new Exception("Could not select upper slot guide point");
            if (!lowerWallPoint.Select4(true, null))
                throw new Exception("Could not select lower slot guide point");
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY - 0.02, 0);
            if (displayDim == null) throw new Exception("Could not create slot guide spacing dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot guide spacing dimension");
            swDim.SystemValue = slotGuideSpacing;
            swModel.ClearSelection2(true);

            line2.Select4(false, null);
            line1.Select4(true, null);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, 0.03, 0);
            if (displayDim == null) throw new Exception("Could not create slot guide angle dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot guide angle dimension");
            swDim.SystemValue = angleInRadians;
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom edge for slot guide offset");
            if (!lowerWallPoint.Select4(true, null))
                throw new Exception("Could not select offset slot guide point");
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY - 0.02, 0);
            if (displayDim == null) throw new Exception("Could not create slot guide offset dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot guide offset dimension");
            swDim.SystemValue = slotGuideOffset;
            swModel.ClearSelection2(true);

            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, topY, 0, false, 0, null, 0);
            swSketchManager.CreateFillet(filletRadius, (int)swConstrainedCornerAction_e.swConstrainedCornerKeepGeometry);
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", -halfWidth, topY, 0, false, 0, null, 0);
            swSketchManager.CreateFillet(filletRadius, (int)swConstrainedCornerAction_e.swConstrainedCornerKeepGeometry);
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, topY - 0.05, 0, false, 0, null, 0);
            swSketchManager.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimClosest, halfWidth, 0.325, 0);

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            SketchSegment lineseg1 = (SketchSegment)swSketchManager.CreateCenterLine(0, 0, 0, 0, bottomY, 0);
            if (lineseg1 == null) throw new Exception("Could not create slot mirror centerline");

            swModel.ClearSelection2(true);

            SelectData mirrorData = selMgr.CreateSelectData();
            if (mirrorData == null) throw new Exception("Could not create slot mirror selection data");
            mirrorData.Mark = 2;

            SelectData mirrorData2 = selMgr.CreateSelectData();
            if (mirrorData2 == null) throw new Exception("Could not create slot entity selection data");
            mirrorData2.Mark = 1;

            slotfilletArc.Select4(false, mirrorData2);
            line1.Select4(true, mirrorData2);
            line2.Select4(true, mirrorData2);
            lineseg1.Select4(true, mirrorData);
            swModel.SketchMirror();
            swModel.ClearSelection2(true);
            lineseg1.Select4(true, null);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);
            lineseg1.Select4(true, null);
            swModel.SketchAddConstraints("sgVERTICAL");
            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", -halfWidth, topY - 0.05, 0, false, 0, null, 0);
            swSketchManager.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimClosest, -halfWidth, 0.325, 0);

            bool status = swModel.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            swModel.SketchManager.CreateConstructionGeometry();

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
            SketchSegment line3 = (SketchSegment)swSketchManager.CreateLine(-halfWidth, bottomY, 0, 0, 0, 0);
            SketchSegment line4 = (SketchSegment)swSketchManager.CreateLine(halfWidth, bottomY, 0, 0, 0, 0);
            line3.Select4(false, null);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            swSketchManager.InsertSketch(true);
            swModel.Extension.SelectByID2("Sketch2", "SKETCH", 0, 0, 0, false, 0, null, 0);
            swModel.FeatureManager.FeatureCut4(false, false, false, (int)swEndConditions_e.swEndCondThroughAll, (int)swEndConditions_e.swEndCondThroughAll, 0, 0, false, false, false, false, 0, 0, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);
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

            swSketchManager.InsertSketch(true);

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
                // Use GetTitle() to ensure we close the specific document we just saved
                _swApp.CloseDoc(swModel.GetTitle());
                Console.WriteLine("Part closed after creating.");
            }

            // Restore user preferences
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
            Console.WriteLine("Done!");

            return Path.GetFileName(savedPath); 
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fatal error: " + ex);
            try { _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true); } catch { }
            return null; // Return null to indicate failure
        }
    }

}





