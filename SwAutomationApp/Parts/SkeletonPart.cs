using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Creates a "skeleton" reference part.
///
/// This part is not normal production geometry.
/// Instead, it contains reference planes that the machine assembly uses as a stable positioning frame.
/// In practice, it behaves like a construction part for assembly layout.
/// </summary>
public sealed class SkeletonPart
{
    
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    public SkeletonPart(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        _pdm = pdm ?? throw new ArgumentNullException(nameof(pdm));
    }

    // File and save settings.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; } = false;
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "skeleton.SLDPRT";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    // Editable reference-plane distances in meters.
    public double SideOffset { get; set; } = 0.5;
    public double GroundOffset { get; set; } = -0.25;

    private string GetRequiredOutputFolder() => AutomationSupport.RequireText(OutputFolder, nameof(OutputFolder), nameof(SkeletonPart));
    private string GetRequiredLocalFileName() => AutomationSupport.RequireText(LocalFileName, nameof(LocalFileName), nameof(SkeletonPart));

    /// <summary>
    /// Creates the reference-plane part and saves it.
    /// </summary>
    public string Create()
    {
        // Read the current object settings once at the start.
        double sideOffset = SideOffset;
        double groundOffset = GroundOffset;
        string outFolder = GetRequiredOutputFolder();
        bool closeAfterCreate = CloseAfterCreate;
        bool saveToPdm = SaveToPdm;

        Directory.CreateDirectory(outFolder);

        // Start from a blank part template.
        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        ModelDoc2 swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

        // Create the reference plane on the positive side.
        swModel.ClearSelection2(true);
        swModel.Extension.SelectByID2("Ebene rechts", "PLANE", 0, 0, 0, false, 0, null, 0);
        Feature sideRight = swModel.FeatureManager.InsertRefPlane(
            (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Distance,
            sideOffset,
            0, 0, 0, 0);
        swModel.ClearSelection2(true);
        sideRight.Name = "NDE_BEARING_CENTER";

        // Create the mirror plane on the opposite side.
        swModel.ClearSelection2(true);
        swModel.Extension.SelectByID2("Ebene rechts", "PLANE", 0, 0, 0, false, 0, null, 0);
        Feature sideLeft = swModel.FeatureManager.InsertRefPlane(264, sideOffset,0, 0, 0, 0)
        ;swModel.ClearSelection2(true);
        sideLeft.Name = "DE_BEARING_CENTER";

        // Create the ground plane.
        // Positive and negative offsets need slightly different SolidWorks handling,
        // so both directions are handled explicitly.
        swModel.ClearSelection2(true);
        swModel.Extension.SelectByID2("Ebene oben", "PLANE", 0, 0, 0, false, 0, null, 0);
        if (groundOffset > 0)
        {
        Feature groundPlane = swModel.FeatureManager.InsertRefPlane(
            8,
            groundOffset,
            0, 0, 0, 0);
        swModel.ClearSelection2(true);
        groundPlane.Name = "Ground_Plane";
        }
        else
        {
            // For negative offsets, flip the direction by selecting the opposite plane and using a positive distance.
            double value = -1 * groundOffset;
            Feature groundPlane = swModel.FeatureManager.InsertRefPlane(
                264,
                value,
                0, 0, 0, 0);
            swModel.ClearSelection2(true);
            groundPlane.Name = "Ground_Plane";
        }   
        
        // Save the reference part either locally or to PDM.
        string savedPath;
        if (saveToPdm)
        {
            savedPath = _pdm.SaveAsPdm(swModel, outFolder, PdmDataCard);
            Console.WriteLine($"Reference-plane part vaulted at: {savedPath}");
        }
        else
        {
            savedPath = Path.Combine(outFolder, GetRequiredLocalFileName());
            swModel.SaveAs3(savedPath, 0, 1);
            Console.WriteLine($"Reference-plane part saved locally at: {savedPath}");
        }

        if (closeAfterCreate)
        {
            // Use GetTitle so the correct open document is closed.
            _swApp.CloseDoc(swModel.GetTitle());
            Console.WriteLine("Part closed after creating.");
        }

        return Path.GetFileName(savedPath);
    }
}




