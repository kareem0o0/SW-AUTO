using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm; // Make sure this matches your namespace
namespace SwAutomation;

public sealed class Part
{
    private readonly SldWorks _swApp;
    private PdmModule _pdm;
    private const double MmToMeters = 0.001;

    public Part(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        _pdm = pdm ?? throw new ArgumentNullException(nameof(pdm));
    }

    public string Create_stator_sheet(string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
    {
        double Mm(double mm) => mm * MmToMeters;

        // Main dimensions (mm) - change these only.
        double outerDiameterMm = 990.0;
        double innerDiameterMm = 640.0;
        double plateThicknessMm = 8.0;
        double slotWidthMm = 15.7;
        double slotBottomYmm = 320.0;
        double slotTopYmm = 405.2;
        double angleInDegrees = 70;
        double filletradius = 1.0;

        // Derived dimensions (m)
        double outerRadius = Mm(outerDiameterMm / 2.0);
        double innerRadius = Mm(innerDiameterMm / 2.0);
        double plateThickness = Mm(plateThicknessMm);
        double halfWidth = Mm(slotWidthMm / 2.0);
        double bottomY = Mm(slotBottomYmm);
        double topY = Mm(slotTopYmm);
        double centerY = (topY + bottomY) / 2.0;
        double leftX = -halfWidth;
        double angleInRadians = angleInDegrees * Math.PI / 180.0;
        double radiusMm = Mm(filletradius);

        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        try
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
            SelectionMgr selMgr = swModel.SelectionManager;

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);

            bool selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select Front Plane");

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCircleByRadius(0, 0, 0, outerRadius);
            swSketchManager.CreateCircleByRadius(0, 0, 0, innerRadius);

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", outerRadius, 0, 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(outerRadius + Mm(20), Mm(20), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = Mm(outerDiameterMm);

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", innerRadius, 0, 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(innerRadius + Mm(20), Mm(20), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = Mm(innerDiameterMm);

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
            Face2 topFace = selMgr.GetSelectedObject6(1, -1) as Face2;
            PartDoc swPart = swModel as PartDoc;
            swPart.ISetEntityName((Entity)topFace, "TopFace");
            // Keep workflow unchanged: ensure top face is selected for the next sketch.
            selected = swModel.Extension.SelectByID2("", "FACE", (outerRadius + innerRadius) / 2.0, 0, plateThickness, false, 0, null, 0);

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCenterRectangle(0, centerY, 0, halfWidth, topY, 0);

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(leftX + Mm(10), bottomY - Mm(10), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = halfWidth * 2;

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, bottomY + Mm(50), 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(halfWidth + Mm(50), bottomY + Mm(50), 0);
            swDim = displayDim.GetDimension();
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

            SketchSegment line1 = (SketchSegment)swSketchManager.CreateLine(Mm(20), Mm(20), 0, Mm(50), Mm(50), 0);
            SketchSegment line2 = (SketchSegment)swSketchManager.CreateLine(Mm(50), Mm(50), 0, Mm(20), Mm(50), 0);

            swModel.Extension.SelectByID2("", "SKETCHPOINT", Mm(20), Mm(20), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, bottomY + Mm(50), 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");

            swModel.Extension.SelectByID2("", "SKETCHPOINT", Mm(20), Mm(50), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, bottomY + Mm(50), 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, Mm(50), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, Mm(20), 0, true, 0, null, 0);
            SketchSegment slotfilletArc = swSketchManager.CreateFillet(radiusMm, (int)swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry);

            line2.Select4(false, null);
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, Mm(50), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, Mm(20), 0, true, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY - Mm(20), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = Mm(5.7);
            swModel.ClearSelection2(true);

            line2.Select4(false, null);
            line1.Select4(true, null);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, Mm(30), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = angleInRadians;
            swModel.ClearSelection2(true);

            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, Mm(44.3), 0, true, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY - Mm(20), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = Mm(-0.7);
            swModel.ClearSelection2(true);

            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, topY, 0, false, 0, null, 0);
            swSketchManager.CreateFillet(radiusMm, (int)swConstrainedCornerAction_e.swConstrainedCornerKeepGeometry);
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", -halfWidth, topY, 0, false, 0, null, 0);
            swSketchManager.CreateFillet(radiusMm, (int)swConstrainedCornerAction_e.swConstrainedCornerKeepGeometry);
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, topY - Mm(50), 0, false, 0, null, 0);
            swSketchManager.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimClosest, halfWidth, Mm(325.0), 0);

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            SketchSegment lineseg1 = (SketchSegment)swSketchManager.CreateCenterLine(0, 0, 0, 0, bottomY, 0);

            swModel.ClearSelection2(true);

            SelectData mirrorData = selMgr.CreateSelectData();
            mirrorData.Mark = 2;

            SelectData mirrorData2 = selMgr.CreateSelectData();
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
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", -halfWidth, topY - Mm(50), 0, false, 0, null, 0);
            swSketchManager.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimClosest, -halfWidth, Mm(325.0), 0);

            bool status = swModel.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            swModel.SketchManager.CreateConstructionGeometry();

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
            SketchSegment line3 = (SketchSegment)swSketchManager.CreateLine(-halfWidth, Mm(320), 0, 0, 0, 0);
            SketchSegment line4 = (SketchSegment)swSketchManager.CreateLine(halfWidth, Mm(320), 0, 0, 0, 0);
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
                60,
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
            if (SaveToPdm)
            {
                savedPath = _pdm.SaveAsPdm(swModel, outFolder);
                Console.WriteLine($"Part saved to PDM: {savedPath}");
            }
            else
            {
                savedPath = Path.Combine(outFolder, "StatorBleche.SLDPRT");
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
            Console.WriteLine("Fatal error: " + ex.Message);
            try { _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true); } catch { }
            return null; // Return null to indicate failure
        }
    }

    public string Create_shaft(string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
    {
        double Mm(double mm) => mm * MmToMeters;

        // Section radii (mm) - editable.
        double radius1Mm = 60.0;
        double radius2Mm = 50.0;
        double radius3Mm = 40.0;
        double radius4Mm = 45.0;
        double radius5Mm = 35.0;

        // Section lengths (mm) - editable.
        double length1Mm = 180.0;
        double length2Mm = 140.0;
        double length3Mm = 220.0;
        double length4Mm = 110.0;
        double length5Mm = 150.0;

        double[] radii = { Mm(radius1Mm), Mm(radius2Mm), Mm(radius3Mm), Mm(radius4Mm), Mm(radius5Mm) };
        double[] lengths = { Mm(length1Mm), Mm(length2Mm), Mm(length3Mm), Mm(length4Mm), Mm(length5Mm) };

        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

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

        SketchManager swSketchManager = swModel.SketchManager;
        double zCursor = 0.0;

        for (int i = 0; i < 5; i++)
        {
            swModel.ClearSelection2(true);

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

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCircleByRadius(0, 0, 0, radii[i]);

            swModel.ClearSelection2(true);
            bool circleSelected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", radii[i], 0, zCursor, false, 0, null, 0);
            if (!circleSelected)
                throw new Exception($"Could not select circle for section {i + 1}.");

            displayDim = (DisplayDimension)swModel.AddDimension2(radii[i] + Mm(20), Mm(20), zCursor);
            if (displayDim == null)
                throw new Exception($"Could not create radius dimension for section {i + 1}.");

            swDim = displayDim.GetDimension();
            if (swDim == null)
                throw new Exception($"Could not access dimension handle for section {i + 1}.");

            // Circle driving sketch dimension is diameter, so use 2 * radius.
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

            zCursor += lengths[i];
        }

        string savedPath;
        if (SaveToPdm)
        {
            savedPath = _pdm.SaveAsPdm(swModel, outFolder);
            Console.WriteLine($"Shaft saved to PDM: {savedPath}");
        }
        else
        {
            savedPath = Path.Combine(outFolder, "shaft.SLDPRT");
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

    public string CreateSkeleton(double sideOffset, double groundOffset, string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
    {
        Directory.CreateDirectory(outFolder);
        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        ModelDoc2 swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

        // Side_Right
        swModel.ClearSelection2(true);
        swModel.Extension.SelectByID2("Ebene rechts", "PLANE", 0, 0, 0, false, 0, null, 0);
        Feature sideRight = swModel.FeatureManager.InsertRefPlane(
            (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Distance,
            sideOffset * MmToMeters,
            0, 0, 0, 0);
        swModel.ClearSelection2(true);
        sideRight.Name = "NDE_BEARING_CENTER";

        // Side_Left (always opposite of Side_Right)
        swModel.ClearSelection2(true);
        swModel.Extension.SelectByID2("Ebene rechts", "PLANE", 0, 0, 0, false, 0, null, 0);
        Feature sideLeft = swModel.FeatureManager.InsertRefPlane(264, sideOffset * MmToMeters,0, 0, 0, 0)
        ;swModel.ClearSelection2(true);
        sideLeft.Name = "DE_BEARING_CENTER";

        // Ground_Plane (signed value as provided)
        swModel.ClearSelection2(true);
        swModel.Extension.SelectByID2("Ebene oben", "PLANE", 0, 0, 0, false, 0, null, 0);
        if (groundOffset > 0)
        {
        Feature groundPlane = swModel.FeatureManager.InsertRefPlane(
            8,
            groundOffset * MmToMeters,
            0, 0, 0, 0);
        swModel.ClearSelection2(true);
        groundPlane.Name = "Ground_Plane";
        }
        else
        {
            // For negative offsets, flip the direction by selecting the opposite plane and using a positive distance.
            double value = -1 *groundOffset * MmToMeters;
            Feature groundPlane = swModel.FeatureManager.InsertRefPlane(
                264,
                value,
                0, 0, 0, 0);
            swModel.ClearSelection2(true);
            groundPlane.Name = "Ground_Plane";
        }   
        

        string savedPath;
        if (SaveToPdm)
        {
            savedPath = _pdm.SaveAsPdm(swModel, outFolder);
            Console.WriteLine($"Reference-plane part vaulted at: {savedPath}");
        }
        else
        {
            savedPath = Path.Combine(outFolder, "skeleton.SLDPRT");
            swModel.SaveAs3(savedPath, 0, 1);
            Console.WriteLine($"Reference-plane part saved locally at: {savedPath}");
        }

        if (closeAfterCreate)
        {
            // Use GetTitle to ensure the specific document is targeted for closing
            _swApp.CloseDoc(swModel.GetTitle());
            Console.WriteLine("Part closed after creating.");
        }

        return Path.GetFileName(savedPath);
    }
    
}
