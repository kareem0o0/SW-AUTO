using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation.TestHarness;

internal sealed class Part
{
    private readonly SldWorks _swApp;
    private const double MmToMeters = 0.001;

    public Part(SldWorks swApp)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
    }

    public string CreateSkeleton(double sideOffset, double groundOffset, string outFolder, bool closeAfterCreate = false)
    {
        Directory.CreateDirectory(outFolder);

        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        ModelDoc2 model = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

        model.ClearSelection2(true);
        model.Extension.SelectByID2("Ebene rechts", "PLANE", 0, 0, 0, false, 0, null, 0);
        Feature sideRight = model.FeatureManager.InsertRefPlane(
            (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Distance,
            sideOffset * MmToMeters,
            0,
            0,
            0,
            0);
        model.ClearSelection2(true);
        sideRight.Name = "NDE_BEARING_CENTER";

        model.Extension.SelectByID2("Ebene rechts", "PLANE", 0, 0, 0, false, 0, null, 0);
        Feature sideLeft = model.FeatureManager.InsertRefPlane(264, sideOffset * MmToMeters, 0, 0, 0, 0);
        model.ClearSelection2(true);
        sideLeft.Name = "DE_BEARING_CENTER";

        model.Extension.SelectByID2("Ebene oben", "PLANE", 0, 0, 0, false, 0, null, 0);
        Feature groundPlane = groundOffset > 0
            ? model.FeatureManager.InsertRefPlane(8, groundOffset * MmToMeters, 0, 0, 0, 0)
            : model.FeatureManager.InsertRefPlane(264, -groundOffset * MmToMeters, 0, 0, 0, 0);
        model.ClearSelection2(true);
        groundPlane.Name = "Ground_Plane";

        string savedPath = Path.Combine(outFolder, "skeleton.SLDPRT");
        model.SaveAs3(savedPath, 0, 1);
        Console.WriteLine($"Reference-plane part saved locally at: {savedPath}");

        if (closeAfterCreate)
        {
            _swApp.CloseDoc(model.GetTitle());
            Console.WriteLine("Part closed after creating.");
        }

        return Path.GetFileName(savedPath);
    }
}
