using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Creates the stator end sheet.
///
/// This part is similar to the stator ring geometry, but it represents the end cap sheet used
/// at the ends of the repeated stator stack.
/// </summary>
public sealed class StatorEndSheetPart
{
    
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    public StatorEndSheetPart(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp;
        _pdm = pdm;
    }

    // File and save settings.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "StatorEndBleche.SLDPRT";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    // Main editable geometry values.
    public double OuterDiameter { get; set; } = 0.99;
    public double InnerDiameter { get; set; } = 0.64;
    public double PlateThickness { get; set; } = 0.001;
    public double SlotWidth { get; set; } = 0.0205;
    public double SlotBottomY { get; set; } = 0.32;
    public double SlotTopY { get; set; } = 0.406;
    public int SlotPatternCount { get; set; } = 60;
    public string MaterialName { get; set; } = "AISI 1020";

    private string GetRequiredOutputFolder() => OutputFolder;
    private string GetRequiredLocalFileName() => LocalFileName;
    private AutomationUiScope BeginAutomationUiSuppression() => new(_swApp);

    /// <summary>
    /// Creates the stator end sheet model and saves it.
    /// </summary>
    public string Create()
    {
        
        using var automationUi = BeginAutomationUiSuppression();

        // Read the current object state first.
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
        double plateThicknessValue = PlateThickness;
        double slotWidth = SlotWidth;
        double slotBottomY = SlotBottomY;
        double slotTopY = SlotTopY;
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

        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        // Build the end sheet and save it once the feature tree is complete.
        {
            Dimension swDim = null;
            DisplayDimension displayDim = null;

            if (!Directory.Exists(outFolder))
                Directory.CreateDirectory(outFolder);

            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            if (swModel == null)
                throw new Exception("Failed to create new part");

            swSketchManager = swModel.SketchManager;

            // Apply the chosen material to the generated part.
            PartDoc statorPart = swModel as PartDoc;
            statorPart.SetMaterialPropertyName2("", "", Name: materialName);

            // Phase 1:
            // Build the base ring as an annulus.
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

            // Phase 2:
            // Create one slot profile, then later pattern it around the part.
            swSketchManager.InsertSketch(true);
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
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY / 2.0, 0);
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

            // Create two helper lines from the slot corners back to the origin.
            // These do not become part of the solid directly; they are support geometry that helps
            // fully constrain the slot location before the cut is created.
            SelectionMgr selectionMgr = swModel.SelectionManager as SelectionMgr;
            if (selectionMgr == null)
                throw new Exception("Could not access selection manager for end-sheet slot references");

            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", -halfWidth, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select left bottom slot corner");
            SketchPoint leftBottomSlotCorner = selectionMgr.GetSelectedObject6(1, -1) as SketchPoint;
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select right bottom slot corner");
            SketchPoint rightBottomSlotCorner = selectionMgr.GetSelectedObject6(1, -1) as SketchPoint;
            swModel.ClearSelection2(true);

            if (leftBottomSlotCorner == null || rightBottomSlotCorner == null)
                throw new Exception("Could not access end-sheet slot corner points");

            SketchSegment line3 = (SketchSegment)swSketchManager.CreateLine(-halfWidth, bottomY, 0, 0, 0, 0);
            SketchSegment line4 = (SketchSegment)swSketchManager.CreateLine(halfWidth, bottomY, 0, 0, 0, 0);
            if (line3 == null || line4 == null)
                throw new Exception("Could not create slot corner lines");

            SketchLine line3Sketch = line3 as SketchLine;
            SketchLine line4Sketch = line4 as SketchLine;
            if (line3Sketch == null || line4Sketch == null)
                throw new Exception("Could not access slot corner line geometry");

            SketchPoint line3Start = line3Sketch.GetStartPoint2();
            SketchPoint line3End = line3Sketch.GetEndPoint2();
            SketchPoint line4Start = line4Sketch.GetStartPoint2();
            SketchPoint line4End = line4Sketch.GetEndPoint2();
            if (line3Start == null || line3End == null || line4Start == null || line4End == null)
                throw new Exception("Could not access slot corner line endpoints");

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

            // Turn the fully defined slot sketch into a through-all cut.
            swSketchManager.InsertSketch(true);
            if (!SelectSketchByIndex(swModel, 2))
                throw new Exception("Could not select slot sketch");
            Feature slotCutFeature = swModel.FeatureManager.FeatureCut4(
                false, false, false,
                (int)swEndConditions_e.swEndCondThroughAll, (int)swEndConditions_e.swEndCondThroughAll,
                0, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                false, true, true, true, true, false,
                0, 0, false, false);
            if (slotCutFeature == null)
                throw new Exception("Failed to create slot cut extrusion");

            // Repeat that one slot evenly around the center axis.
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
                throw new Exception("Failed to create circular slot pattern");

            // Save the completed part by the selected workflow.
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

            Console.WriteLine("Done!");

            return Path.GetFileName(savedPath);
        }
    }

}





