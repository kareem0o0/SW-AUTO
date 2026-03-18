using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Creates a stepped shaft by building several circular sections one after another.
///
/// Each section is defined by:
/// - one radius
/// - one length
///
/// The method loops through the sections, sketches a circle, and extrudes that section.
/// </summary>
public sealed class ShaftPart
{
    
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    public ShaftPart(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp;
        _pdm = pdm;
    }

    // File and save settings.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "shaft.SLDPRT";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    // Shaft step radii.
    public double Radius1 { get; set; } = 0.06;
    public double Radius2 { get; set; } = 0.05;
    public double Radius3 { get; set; } = 0.04;
    public double Radius4 { get; set; } = 0.045;
    public double Radius5 { get; set; } = 0.035;

    // Shaft step lengths.
    public double Length1 { get; set; } = 0.18;
    public double Length2 { get; set; } = 0.14;
    public double Length3 { get; set; } = 0.22;
    public double Length4 { get; set; } = 0.11;
    public double Length5 { get; set; } = 0.15;
    public string MaterialName { get; set; } = "AISI 1020";

    private string GetRequiredOutputFolder() => OutputFolder;
    private string GetRequiredLocalFileName() => LocalFileName;
    private AutomationUiScope BeginAutomationUiSuppression() => new(_swApp);

    /// <summary>
    /// Creates the shaft geometry and saves it.
    /// </summary>
    public string Create()
    {
        
        using var automationUi = BeginAutomationUiSuppression();

        // Read the current object state into local variables for clarity.
        string outFolder = GetRequiredOutputFolder();
        bool closeAfterCreate = CloseAfterCreate;
        bool saveToPdm = SaveToPdm;
        string materialName = MaterialName;

        // Editable section radii in meters.
        double radius1 = Radius1;
        double radius2 = Radius2;
        double radius3 = Radius3;
        double radius4 = Radius4;
        double radius5 = Radius5;

        // Editable section lengths in meters.
        double length1 = Length1;
        double length2 = Length2;
        double length3 = Length3;
        double length4 = Length4;
        double length5 = Length5;

        double[] radii = { radius1, radius2, radius3, radius4, radius5 };
        double[] lengths = { length1, length2, length3, length4, length5 };

        Directory.CreateDirectory(outFolder);

        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        ModelDoc2 swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
        if (swModel == null)
            throw new Exception("Failed to create new part document.");

        Dimension swDim = null;
        DisplayDimension displayDim = null;

        bool SelectSketchByIndex(int index)
        {
            return swModel.Extension.SelectByID2($"Skizze{index}", "SKETCH", 0, 0, 0, false, 0, null, 0)
                || swModel.Extension.SelectByID2($"Sketch{index}", "SKETCH", 0, 0, 0, false, 0, null, 0);
        }

        // zCursor keeps track of the end of the previous shaft section.
        // Each new section is built starting from that current end face.
        SketchManager swSketchManager = swModel.SketchManager;
        double zCursor = 0.0;

        for (int i = 0; i < 5; i++)
        {
            swModel.ClearSelection2(true);

            // The first sketch is drawn on the front plane.
            // Later sketches are drawn on the face created by the previous extrusion.
            bool sketchPlaneSelected;
            if (i == 0)
            {
                sketchPlaneSelected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            }
            else
            {
                sketchPlaneSelected = swModel.Extension.SelectByID2("", "FACE", 0, 0, zCursor, false, 0, null, 0);
            }

            if (!sketchPlaneSelected)
                throw new Exception($"Could not select sketch plane/face for shaft section {i + 1}.");

            // Sketch the section circle.
            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCircleByRadius(0, 0, 0, radii[i]);

            swModel.ClearSelection2(true);
            bool circleSelected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", radii[i], 0, zCursor, false, 0, null, 0);
            if (!circleSelected)
                throw new Exception($"Could not select circle for section {i + 1}.");

            displayDim = (DisplayDimension)swModel.AddDimension2(radii[i] + 0.02, 0.02, zCursor);
            if (displayDim == null)
                throw new Exception($"Could not create radius dimension for section {i + 1}.");

            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception($"Could not access dimension handle for section {i + 1}.");

            // SolidWorks stores the sketch circle driving dimension as diameter, not radius.
            swDim.SystemValue = radii[i] * 2.0;

            swSketchManager.InsertSketch(true);

            swModel.ClearSelection2(true);
            bool sketchSelected = SelectSketchByIndex(i + 1);
            if (!sketchSelected)
                throw new Exception($"Could not select sketch for section {i + 1}.");

            bool featureCreated = swModel.FeatureManager.FeatureExtrusion2(
                true, false, false,
                (int)swEndConditions_e.swEndCondBlind, 0,
                lengths[i], 0,
                false, false, false, false,
                0, 0, false, false, false, false,
                true, true, true, 0, 0, false) != null;

            if (!featureCreated)
                throw new Exception($"Failed to create shaft section {i + 1}.");

            // Move the cursor forward so the next section starts where this one ends.
            zCursor += lengths[i];
        }

        PartDoc shaftPart = swModel as PartDoc;
        // Apply material after the geometry exists.
        shaftPart.SetMaterialPropertyName2("", "", Name: materialName);

        // Save either locally or to PDM.
        string savedPath;
        if (saveToPdm)
        {
            savedPath = _pdm.SaveAsPdm(swModel, outFolder, PdmDataCard);
            Console.WriteLine($"Shaft saved to PDM: {savedPath}");
        }
        else
        {
            savedPath = Path.Combine(outFolder, GetRequiredLocalFileName());
            swModel.SaveAs3(savedPath, 0, 1);
            Console.WriteLine($"Shaft saved locally: {savedPath}");
        }

        if (closeAfterCreate)
        {
            _swApp.CloseDoc(swModel.GetTitle());
            Console.WriteLine("Part closed after creating.");
        }

        return Path.GetFileName(savedPath);
    }
}





