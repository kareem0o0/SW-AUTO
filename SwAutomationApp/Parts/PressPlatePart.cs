using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Creates the press plate.
///
/// This part is made from:
/// - a ring body
/// - one plate segment
/// - a circular pattern that repeats the plate segment around the ring
///
/// It also stores AssemblyAngleDeg because that angle is needed later when the part is placed
/// in the machine assembly.
/// </summary>
public sealed class PressPlatePart
{
    
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    public PressPlatePart(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp;
        _pdm = pdm;
    }

    // File and save settings.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "PressPlate.SLDPRT";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    // Main editable geometry values.
    public double OuterDiameter { get; set; } = 0.99;
    public double RingInnerDiameter { get; set; } = 0.84;
    public double PlateOuterInsetFromOuterDiameter { get; set; } = 0.005;
    public double PlateRadialLength { get; set; } = 0.165;
    public double RingThickness { get; set; } = 0.002;
    public double PlateBodyThickness { get; set; } = 0.01;
    public double PlateWidth { get; set; } = 0.006;
    public int PlateCount { get; set; } = 60;
    public double AssemblyAngleDeg { get; set; } = 3.0;
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
    /// Creates the press plate model and saves it.
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
        double ringInnerDiameter = RingInnerDiameter;
        double plateOuterInsetFromOuterDiameterValue = PlateOuterInsetFromOuterDiameter;
        double plateRadialLengthValue = PlateRadialLength;
        double ringThicknessValue = RingThickness;
        double plateBodyThicknessValue = PlateBodyThickness;
        double plateWidthValue = PlateWidth;
        int plateCount = PlateCount;
        string materialName = MaterialName;

        // Derived dimensions
        double outerRadius = outerDiameter / 2.0;
        double ringInnerRadius = ringInnerDiameter / 2.0;
        double plateOuterInsetFromOuterDiameter = plateOuterInsetFromOuterDiameterValue;
        double plateRadialLength = plateRadialLengthValue;
        double ringThickness = ringThicknessValue;
        double plateBodyThickness = plateBodyThicknessValue;
        double plateHalfWidth = plateWidthValue / 2.0;
        double plateOuterY = outerRadius - plateOuterInsetFromOuterDiameter;
        double plateInnerY = plateOuterY - plateRadialLength;
        double plateCenterY = (plateOuterY + plateInnerY) / 2.0;
        double plateHeight = plateRadialLength;

        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        // Build the ring, add one plate, pattern it around the body, then save it.
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

            // Apply the chosen material.
            PartDoc pressPlatePart = swModel as PartDoc;
            pressPlatePart.SetMaterialPropertyName2("", "", Name: materialName);

            // Phase 1:
            // Build the annular ring.
            bool selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select Front Plane for press plate");

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCircleByRadius(0, 0, 0, outerRadius);
            swSketchManager.CreateCircleByRadius(0, 0, 0, ringInnerRadius);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", outerRadius, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select press-plate outer circle");
            displayDim = (DisplayDimension)swModel.AddDimension2(outerRadius + 0.02, 0.02, 0);
            if (displayDim == null) throw new Exception("Could not create press-plate outer diameter dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate outer diameter dimension");
            swDim.SystemValue = outerDiameter;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", ringInnerRadius, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select press-plate ring inner circle");
            displayDim = (DisplayDimension)swModel.AddDimension2(ringInnerRadius + 0.02, 0.02, 0);
            if (displayDim == null) throw new Exception("Could not create press-plate ring inner diameter dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate ring inner diameter dimension");
            swDim.SystemValue = ringInnerDiameter;

            swModel.ClearSelection2(true);
            if (!SelectSketchByIndex(swModel, 1))
                throw new Exception("Could not select press-plate ring sketch");
            Feature ringFeature = swModel.FeatureManager.FeatureExtrusion2(
                true, false, false,
                (int)swEndConditions_e.swEndCondBlind, 0,
                ringThickness, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                true, true, true, 0, 0, false);
            if (ringFeature == null)
                throw new Exception("Failed to create press-plate ring");

            // Phase 2:
            // Sketch one rectangular plate segment and extrude it as a separate body/feature.
            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not reselect Front Plane for press-plate body");

            swSketchManager.InsertSketch(true);
            SketchSegment topLineSegment = swSketchManager.CreateLine(-plateHalfWidth, plateOuterY, 0, plateHalfWidth, plateOuterY, 0) as SketchSegment;
            SketchSegment rightLineSegment = swSketchManager.CreateLine(plateHalfWidth, plateOuterY, 0, plateHalfWidth, plateInnerY, 0) as SketchSegment;
            SketchSegment bottomLineSegment = swSketchManager.CreateLine(plateHalfWidth, plateInnerY, 0, -plateHalfWidth, plateInnerY, 0) as SketchSegment;
            SketchSegment leftLineSegment = swSketchManager.CreateLine(-plateHalfWidth, plateInnerY, 0, -plateHalfWidth, plateOuterY, 0) as SketchSegment;
            SketchLine topLine = topLineSegment as SketchLine;
            SketchLine rightLine = rightLineSegment as SketchLine;
            SketchLine bottomLine = bottomLineSegment as SketchLine;
            SketchLine leftLine = leftLineSegment as SketchLine;
            if (topLineSegment == null || rightLineSegment == null || bottomLineSegment == null || leftLineSegment == null
                || topLine == null || rightLine == null || bottomLine == null || leftLine == null)
                throw new Exception("Could not create press-plate body rectangle");

            SketchPoint LineStart(SketchLine line)
            {
                return line.GetStartPoint2();
            }

            SketchPoint LineEnd(SketchLine line)
            {
                return line.GetEndPoint2();
            }

            // First connect the four loose lines into one closed rectangular profile.
            // SolidWorks needs these endpoint relations before it can treat the shape as a valid
            // extrusion profile.
            if (!(LineEnd(topLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate top-right point for closure");
            if (!(LineStart(rightLine)?.Select4(true, null) ?? false))
                throw new Exception("Could not select press-plate right-top point for closure");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!(LineEnd(rightLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate bottom-right point for closure");
            if (!(LineStart(bottomLine)?.Select4(true, null) ?? false))
                throw new Exception("Could not select press-plate bottom-right bottom-line point for closure");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!(LineEnd(bottomLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate bottom-left point for closure");
            if (!(LineStart(leftLine)?.Select4(true, null) ?? false))
                throw new Exception("Could not select press-plate left-bottom point for closure");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!(LineEnd(leftLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate left-top point for closure");
            if (!(LineStart(topLine)?.Select4(true, null) ?? false))
                throw new Exception("Could not select press-plate top-left point for closure");
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!topLineSegment.Select4(false, null))
                throw new Exception("Could not select press-plate top line");
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            if (!bottomLineSegment.Select4(false, null))
                throw new Exception("Could not select press-plate bottom line");
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            if (!rightLineSegment.Select4(false, null))
                throw new Exception("Could not select press-plate right line");
            swModel.SketchAddConstraints("sgVERTICAL");
            swModel.ClearSelection2(true);

            if (!leftLineSegment.Select4(false, null))
                throw new Exception("Could not select press-plate left line");
            swModel.SketchAddConstraints("sgVERTICAL");
            swModel.ClearSelection2(true);

            // Add the driving dimensions that define this one plate segment.
            // These values control the segment's width, height, centering, and radial placement.
            swModel.ClearSelection2(true);
            if (!(LineStart(topLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate body top-left point");
            if (!(LineEnd(topLine)?.Select4(true, null) ?? false))
                throw new Exception("Could not select press-plate body top-right point for width");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(0.01, plateOuterY + 0.01, 0);
            if (displayDim == null) throw new Exception("Could not create press-plate body width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate body width dimension");
            swDim.SystemValue = plateHalfWidth * 2.0;

            swModel.ClearSelection2(true);
            if (!(LineEnd(rightLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate body bottom-right point");
            if (!(LineStart(rightLine)?.Select4(true, null) ?? false))
                throw new Exception("Could not select press-plate body top-right point");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(plateHalfWidth + 0.02, plateCenterY, 0);
            if (displayDim == null) throw new Exception("Could not create press-plate body height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate body height dimension");
            swDim.SystemValue = plateHeight;

            swModel.ClearSelection2(true);
            if (!(LineEnd(topLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate body top-right point for horizontal location");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for press-plate body horizontal location");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(plateHalfWidth / 2.0, plateOuterY + 0.02, 0);
            if (displayDim == null) throw new Exception("Could not create press-plate body horizontal location dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate body horizontal location dimension");
            swDim.SystemValue = plateHalfWidth;
            swModel.ClearSelection2(true);

            if (!(LineEnd(topLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not reselect press-plate body top-right point");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for press-plate body top offset");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(plateHalfWidth + 0.02, plateOuterY / 2.0, 0);
            if (displayDim == null) throw new Exception("Could not create press-plate body top offset dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate body top offset dimension");
            swDim.SystemValue = plateOuterY;
            swModel.ClearSelection2(true);

            swSketchManager.InsertSketch(true);
            if (!SelectSketchByIndex(swModel, 2))
                throw new Exception("Could not select press-plate body sketch");
            Feature plateFeature = swModel.FeatureManager.FeatureExtrusion2(
                true, true, false,
                (int)swEndConditions_e.swEndCondBlind, 0,
                plateBodyThickness, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                false, true, true, 0, 0, false);
            if (plateFeature == null)
                throw new Exception("Failed to create separate press-plate body");

            // Phase 3:
            // Pattern that one plate body around the ring.
            // This keeps the seed feature simple: we only model one segment, then repeat it.
            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("Z-Achse", "AXIS", 0, 0, 0, false, 1, null, 0);
            if (!selected) throw new Exception("Could not select Z axis for press-plate circular pattern");
            if (!plateFeature.Select2(true, 4))
                throw new Exception("Could not select press-plate body feature for circular pattern");
            Feature platePattern = (Feature)swModel.FeatureManager.FeatureCircularPattern5(
                plateCount,
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
            if (platePattern == null)
                throw new Exception("Failed to create press-plate circular pattern");

            swModel.ClearSelection2(true);
            swModel.ShowNamedView2("*Front", (int)swStandardViews_e.swFrontView);
            swModel.ViewZoomtofit2();

            // Save either as a normal disk file or as a controlled PDM file.
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





