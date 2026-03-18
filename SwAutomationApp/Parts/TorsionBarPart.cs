using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Creates the torsion bar part and also acts as the data owner for the torsion-bar drawing.
///
/// This class is a good example of the project's architecture:
/// - the part object owns the 3D settings
/// - the same part object also owns the drawing settings
/// - the detailed drawing logic itself stays in drawing.cs
/// </summary>
public sealed class TorsionBarPart
{
    
    private const string DefaultDrawingTemplateFolderPath = @"C:\Users\kareem.salah\PDM\Birr Machines PDM\40_Templates\Solidworks\Blattformate\Birr Machines";
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    public TorsionBarPart(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp;
        _pdm = pdm;
    }

    // File and save settings for the 3D part.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "TorsionBar.SLDPRT";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    // Editable part geometry and configuration settings.
    public double BarLength { get; set; } = 1.074;
    public double BarHeight { get; set; } = 0.04;
    public double BarThickness { get; set; } = 0.03;
    public double HoleCenterlineOffsetFromBottom { get; set; } = 0.02;
    public double OuterHoleEndOffset { get; set; } = 0.03;
    public double HolePairSpacing { get; set; } = 0.315;
    public double OuterHoleDiameter { get; set; } = 0.01;
    public double InnerHoleDiameter { get; set; } = 0.016;
    public double CenterHoleDiameter { get; set; } = 0.016;
    public string OuterTapSizePrimary { get; set; } = "M10x1.5";
    public string OuterTapSizeFallback { get; set; } = "M10";
    public string InnerTapSizePrimary { get; set; } = "M16x2";
    public string InnerTapSizeFallback { get; set; } = "M16";
    public string P0001ConfigName { get; set; } = "P0001";
    public string P0002ConfigName { get; set; } = "P0002";
    public string MaterialName { get; set; } = "AISI 1020";

    // Drawing settings stay on the part object for the same reason the 3D settings do:
    // this class is the single source of truth for everything related to the torsion bar.
    // drawing.cs will read these values, but it does not own them.
    public string DrawingOutputFolder { get; set; } = string.Empty;
    public bool DrawingCloseAfterCreate { get; set; }
    public bool DrawingSaveToPdm { get; set; }
    public string DrawingLocalFileName { get; set; } = "TorsionBar.SLDDRW";
    public BirrDataCardValues DrawingPdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();
    public string DrawingSheetName { get; set; } = "Torsion Bar";
    public string DrawingLanguageCode { get; set; } = "EN";
    public string DrawingTemplateFolderPath { get; set; } = DefaultDrawingTemplateFolderPath;
    public bool DrawingPreferSolidWorksTemplateLocations { get; set; } = true;
    public string DrawingSheetFormatPathOverride { get; set; } = string.Empty;
    public string DrawingTemplatePathOverride { get; set; } = string.Empty;
    public bool DrawingUseFirstAngleProjection { get; set; }
    public double DrawingBottomTitleBlockClearance { get; set; } = 0.085;
    public string DrawingReferencedConfiguration { get; set; } = "P0002";

    private string GetRequiredOutputFolder() => OutputFolder;
    private string GetRequiredLocalFileName() => LocalFileName;
    private AutomationUiScope BeginAutomationUiSuppression() => new(_swApp);

    // Public creation choices:
    // CreatePart() makes only the 3D model.
    // Create() makes the 3D model and then the drawing.
    public string Create()
    {
        return DrawingMethods.CreateTorsionBarDrawing(this, _swApp, _pdm);
    }

    /// <summary>
    /// Creates only the 3D torsion bar part.
    /// </summary>
    public string CreatePart()
    {
        
        using var automationUi = BeginAutomationUiSuppression();

        // Read the current object state first so the whole build uses one stable set of values.
        string outFolder = GetRequiredOutputFolder();
        bool closeAfterCreate = CloseAfterCreate;
        bool saveToPdm = SaveToPdm;

        // SolidWorks can name sketches in German or English depending on the installation.
        // This helper lets the rest of the method ask for "the second sketch" without caring
        // which language the UI used when the sketch was created.
        bool SelectSketchByIndex(ModelDoc2 model, int index)
        {
            return model.Extension.SelectByID2($"Skizze{index}", "SKETCH", 0, 0, 0, false, 0, null, 0)
                || model.Extension.SelectByID2($"Sketch{index}", "SKETCH", 0, 0, 0, false, 0, null, 0);
        }

        // Main dimensions (m) - change these only.
        double barLengthValue = BarLength;
        double barHeightValue = BarHeight;
        double barThicknessValue = BarThickness;
        double holeCenterlineOffsetFromBottom = HoleCenterlineOffsetFromBottom;
        double outerHoleEndOffset = OuterHoleEndOffset;
        double holePairSpacing = HolePairSpacing;
        double outerHoleDiameterValue = OuterHoleDiameter;
        double innerHoleDiameterValue = InnerHoleDiameter;
        double centerHoleDiameterValue = CenterHoleDiameter;
        string outerTapSizePrimary = OuterTapSizePrimary;
        string outerTapSizeFallback = OuterTapSizeFallback;
        string innerTapSizePrimary = InnerTapSizePrimary;
        string innerTapSizeFallback = InnerTapSizeFallback;

        string p0001ConfigName = P0001ConfigName;
        string p0002ConfigName = P0002ConfigName;
        string materialName = MaterialName;

        // Derived dimensions.
        // The bar is modeled around the origin, so half sizes and centered offsets make the
        // later sketching and hole placement steps much easier to read.
        double barLength = barLengthValue;
        double barHeight = barHeightValue;
        double barThickness = barThicknessValue;
        double halfLength = barLength / 2.0;
        double halfHeight = barHeight / 2.0;
        double halfThickness = barThickness / 2.0;
        double holeCenterlineY = -halfHeight + holeCenterlineOffsetFromBottom;
        double outerHoleCenterX = halfLength - outerHoleEndOffset;
        double innerHoleCenterX = outerHoleCenterX - holePairSpacing;
        double centerHoleRadius = centerHoleDiameterValue / 2.0;

        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        // Build the torsion bar from the current property values, then save it.
        {
            Dimension swDim = null;
            DisplayDimension displayDim = null;

            if (!Directory.Exists(outFolder))
                Directory.CreateDirectory(outFolder);

            // Start from a new part template and activate it explicitly.
            // This extra activation helps avoid first-run SolidWorks issues.
            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            if (swModel == null)
                throw new Exception("Failed to create new part");

            int activateDocError = 0;
            ModelDoc2 activePartModel = (ModelDoc2)_swApp.ActivateDoc3(
                swModel.GetTitle(),
                true,
                (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
                ref activateDocError);
            if (activePartModel == null)
                throw new Exception($"Could not activate the torsion-bar part document. ActivateDoc3 error code: {activateDocError}");

            swSketchManager = swModel.SketchManager;

            // The material is applied early so the saved file already carries the correct metadata.
            PartDoc torsionBarPart = swModel as PartDoc;
            torsionBarPart.SetMaterialPropertyName2("", "", Name: materialName);

            ConfigurationManager configMgr = swModel.ConfigurationManager;
            if (configMgr == null)
                throw new Exception("Could not access configuration manager");

            Configuration p0001Config = configMgr.ActiveConfiguration;
            if (p0001Config == null)
                throw new Exception("Could not access the active torsion-bar configuration");
            p0001Config.Name = p0001ConfigName;

            // Phase 1:
            // Build the main bar body as a centered rectangle extruded with a mid-plane thickness.
            bool selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select Front Plane for torsion bar");

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCenterRectangle(0, 0, 0, halfLength, halfHeight, 0);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, halfHeight, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select torsion-bar top edge");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(0, halfHeight + 0.02, 0);
            if (displayDim == null) throw new Exception("Could not create torsion-bar length dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access torsion-bar length dimension");
            swDim.SystemValue = barLength;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfLength, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select torsion-bar right edge");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(halfLength + 0.02, 0, 0);
            if (displayDim == null) throw new Exception("Could not create torsion-bar height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access torsion-bar height dimension");
            swDim.SystemValue = barHeight;

            swModel.ClearSelection2(true);
            if (!SelectSketchByIndex(swModel, 1))
                throw new Exception("Could not select torsion-bar base sketch");
            Feature baseExtrude = (Feature)swModel.FeatureManager.FeatureExtrusion2(
                true, false, true,
                (int)swEndConditions_e.swEndCondMidPlane, 0,
                barThickness, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                true, true, true, 0, 0, false);
            if (baseExtrude == null)
                throw new Exception("Failed to create torsion-bar base extrusion");

            swModel.ClearSelection2(true);
            swModel.EditRebuild3();

            // Phase 2:
            // Create the tapped hole wizard features that belong only to the second configuration.
            Feature CreateTappedHoleFeature(string[] sizeCandidates, double nominalDiameter, double centerX, string holeLabel)
            {
                foreach (string sizeCandidate in sizeCandidates)
                {
                    swModel.ClearSelection2(true);
                    bool faceSelected = swModel.Extension.SelectByID2("", "FACE", centerX, holeCenterlineY, halfThickness, false, 0, null, 0);
                    if (!faceSelected)
                        throw new Exception($"Could not select torsion-bar front face for {holeLabel}");

                    Feature tappedHoleFeature = swModel.FeatureManager.HoleWizard5(
                        (int)swWzdGeneralHoleTypes_e.swWzdTap,
                        (int)swWzdHoleStandards_e.swStandardISO,
                        (int)swWzdHoleStandardFastenerTypes_e.swStandardISOTappedHole,
                        sizeCandidate,
                        (short)swEndConditions_e.swEndCondThroughAll,
                        nominalDiameter,
                        barThickness,
                        0,
                        -1, -1, -1, -1, -1, -1, -1, -1,
                        (double)(int)swWzdHoleCosmeticThreadTypes_e.swCosmeticThreadWithoutCallout,
                        (double)(int)swWzdHoleThreadEndCondition_e.swEndThreadTypeTHROUGH_ALL,
                        (double)(int)swWzdHoleHcoilTapTypes_e.swTapTypePlug,
                        0,
                        "",
                        false, false, true, false, true, false);
                    if (tappedHoleFeature != null)
                        return tappedHoleFeature;
                }

                throw new Exception($"Failed to create {holeLabel} hole-wizard feature");
            }

            object centerHole = null;
            for (int attempt = 0; attempt < 2 && centerHole == null; attempt++)
            {
                // The center-hole sketch is retried because fresh SolidWorks sessions can
                // occasionally fail to enter the face sketch on the first attempt.
                swModel.ClearSelection2(true);
                selected = swModel.Extension.SelectByID2("", "FACE", 0, 0, halfThickness, false, 0, null, 0);
                if (!selected)
                    throw new Exception("Could not select torsion-bar front face for center-hole sketch");

                swModel.InsertSketch2(true);

                if (swModel.GetActiveSketch2() == null)
                {
                    swModel.EditRebuild3();
                    continue;
                }

                centerHole = swSketchManager.CreateCircleByRadius(0, holeCenterlineY, 0, centerHoleRadius);
                if (centerHole == null)
                    centerHole = swModel.CreateCircleByRadius2(0, holeCenterlineY, 0, centerHoleRadius);

                if (centerHole == null)
                {
                    swModel.InsertSketch2(true);
                    swModel.EditRebuild3();
                }
            }

            if (centerHole == null)
                throw new Exception("Could not create torsion-bar center-hole sketch");

            // Turn the center sketch into a through cut.
            swModel.InsertSketch2(true);
            if (!SelectSketchByIndex(swModel, 2))
                throw new Exception("Could not select torsion-bar center-hole sketch");
            Feature centerHoleCutFeature = swModel.FeatureManager.FeatureCut4(
                false, false, false,
                (int)swEndConditions_e.swEndCondThroughAll, (int)swEndConditions_e.swEndCondThroughAll,
                0, 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                false, true, true, true, true, false,
                0, 0, false, false);
            if (centerHoleCutFeature == null)
                throw new Exception("Failed to create torsion-bar center-hole cut");

            Feature leftOuterTappedHoleFeature = CreateTappedHoleFeature(
                new[] { outerTapSizePrimary, outerTapSizeFallback },
                outerHoleDiameterValue,
                -outerHoleCenterX,
                "left outer M10");
            Feature rightOuterTappedHoleFeature = CreateTappedHoleFeature(
                new[] { outerTapSizePrimary, outerTapSizeFallback },
                outerHoleDiameterValue,
                outerHoleCenterX,
                "right outer M10");
            Feature leftInnerTappedHoleFeature = CreateTappedHoleFeature(
                new[] { innerTapSizePrimary, innerTapSizeFallback },
                innerHoleDiameterValue,
                -innerHoleCenterX,
                "left inner M16");
            Feature rightInnerTappedHoleFeature = CreateTappedHoleFeature(
                new[] { innerTapSizePrimary, innerTapSizeFallback },
                innerHoleDiameterValue,
                innerHoleCenterX,
                "right inner M16");

            // Phase 3:
            // Create the second configuration and suppress the extra holes in P0001
            // so P0001 and P0002 can represent two manufacturing states of the same part.
            Configuration p0002Config = configMgr.AddConfiguration2(
                p0002ConfigName,
                "",
                "",
                0,
                "",
                "",
                true);
            if (p0002Config == null)
                throw new Exception("Could not create torsion-bar P0002 configuration");

            object p0001Configs = new string[] { p0001ConfigName };
            Feature[] p0002OnlyFeatures =
            {
                centerHoleCutFeature,
                leftOuterTappedHoleFeature,
                rightOuterTappedHoleFeature,
                leftInnerTappedHoleFeature,
                rightInnerTappedHoleFeature
            };

            foreach (Feature p0002OnlyFeature in p0002OnlyFeatures)
            {
                if (p0002OnlyFeature == null)
                    throw new Exception("Could not access a torsion-bar feature for configuration suppression");

                if (!p0002OnlyFeature.SetSuppression2(
                (int)swFeatureSuppressionAction_e.swSuppressFeature,
                (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                p0001Configs))
                {
                    throw new Exception("Could not suppress a torsion-bar feature in P0001");
                }
            }

            object p0002Configs = new string[] { p0002ConfigName };
            foreach (Feature p0002OnlyFeature in p0002OnlyFeatures)
            {
                if (!p0002OnlyFeature.SetSuppression2(
                (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                (int)swInConfigurationOpts_e.swSpecifyConfiguration,
                p0002Configs))
                {
                    throw new Exception("Could not unsuppress a torsion-bar feature in P0002");
                }
            }

            swModel.EditRebuild3();
            swModel.ShowNamedView2("*Front", (int)swStandardViews_e.swFrontView);
            swModel.ViewZoomtofit2();

            // Save either to PDM or to a local file.
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





