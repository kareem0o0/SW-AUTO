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

    private AutomationUiScope BeginAutomationUiSuppression()
    {
        return new AutomationUiScope(_swApp);
    }

    private sealed class AutomationUiScope : IDisposable
    {
        private readonly SldWorks _swApp;
        private readonly bool _originalCommandInProgress;
        private readonly bool _originalInputDimValOnCreate;
        private readonly bool _originalEnableConfirmationCorner;
        private readonly bool _originalSketchPreviewDimensionOnSelect;

        public AutomationUiScope(SldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _originalCommandInProgress = _swApp.CommandInProgress;
            _originalInputDimValOnCreate = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _originalEnableConfirmationCorner = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner);
            _originalSketchPreviewDimensionOnSelect = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect);

            _swApp.CommandInProgress = true;
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner, false);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect, false);
        }

        public void Dispose()
        {
            try
            {
                _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, _originalInputDimValOnCreate);
                _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner, _originalEnableConfirmationCorner);
                _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect, _originalSketchPreviewDimensionOnSelect);
            }
            finally
            {
                _swApp.CommandInProgress = _originalCommandInProgress;
            }
        }
    }

    public string Create_stator_sheet(string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
    {
        double Mm(double mm) => mm * MmToMeters;
        using var automationUi = BeginAutomationUiSuppression();

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

           
            

            // Apply material to the  part
            PartDoc statorPart = swModel as PartDoc;
            statorPart.SetMaterialPropertyName2("", "", Name: "AISI 1020");

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

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCenterRectangle(0, centerY, 0, halfWidth, topY, 0);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom edge");
            displayDim = (DisplayDimension)swModel.AddDimension2(leftX + Mm(10), bottomY - Mm(10), 0);
            if (displayDim == null) throw new Exception("Could not create slot width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot width dimension");
            swDim.SystemValue = halfWidth * 2;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom-right point");
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, topY, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot top-right point");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(halfWidth + Mm(50), centerY, 0);
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

            SketchSegment line1 = (SketchSegment)swSketchManager.CreateLine(Mm(20), Mm(20), 0, Mm(50), Mm(50), 0);
            SketchSegment line2 = (SketchSegment)swSketchManager.CreateLine(Mm(50), Mm(50), 0, Mm(20), Mm(50), 0);
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

            swModel.Extension.SelectByID2("", "SKETCHPOINT", Mm(20), Mm(20), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, bottomY + Mm(50), 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");

            swModel.Extension.SelectByID2("", "SKETCHPOINT", Mm(20), Mm(50), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, bottomY + Mm(50), 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            if (!line2.Select4(false, null))
                throw new Exception("Could not select upper slot guide segment for fillet");
            if (!line1.Select4(true, null))
                throw new Exception("Could not select lower slot guide segment for fillet");
            SketchSegment slotfilletArc = swSketchManager.CreateFillet(radiusMm, (int)swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry);
            if (slotfilletArc == null) throw new Exception("Could not create slot fillet");

            line2.Select4(false, null);
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            if (!upperWallPoint.Select4(false, null))
                throw new Exception("Could not select upper slot guide point");
            if (!lowerWallPoint.Select4(true, null))
                throw new Exception("Could not select lower slot guide point");
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY - Mm(20), 0);
            if (displayDim == null) throw new Exception("Could not create slot guide spacing dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot guide spacing dimension");
            swDim.SystemValue = Mm(5.7);
            swModel.ClearSelection2(true);

            line2.Select4(false, null);
            line1.Select4(true, null);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, Mm(30), 0);
            if (displayDim == null) throw new Exception("Could not create slot guide angle dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot guide angle dimension");
            swDim.SystemValue = angleInRadians;
            swModel.ClearSelection2(true);

            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom edge for slot guide offset");
            if (!lowerWallPoint.Select4(true, null))
                throw new Exception("Could not select offset slot guide point");
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY - Mm(20), 0);
            if (displayDim == null) throw new Exception("Could not create slot guide offset dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot guide offset dimension");
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
            Console.WriteLine("Fatal error: " + ex);
            try { _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true); } catch { }
            return null; // Return null to indicate failure
        }
    }

    public string Create_shaft(string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
    {
        double Mm(double mm) => mm * MmToMeters;
        using var automationUi = BeginAutomationUiSuppression();


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

        PartDoc shaftPart = swModel as PartDoc;
        // Apply material to the  part
        shaftPart.SetMaterialPropertyName2("", "", Name: "AISI 1020");

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
    public string Create_stator_distance_sheet(string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
    {
        double Mm(double mm) => mm * MmToMeters;
        using var automationUi = BeginAutomationUiSuppression();
        bool SelectSketchByIndex(ModelDoc2 model, int index)
        {
            return model.Extension.SelectByID2($"Skizze{index}", "SKETCH", 0, 0, 0, false, 0, null, 0)
                || model.Extension.SelectByID2($"Sketch{index}", "SKETCH", 0, 0, 0, false, 0, null, 0);
        }

        // Main dimensions (mm) - change these only.
        double outerDiameterMm = 990.0;
        double innerDiameterMm = 640.0;
        double plateThicknessMm = 1.0;
        double slotWidthMm = 20.5;
        double slotBottomYmm = innerDiameterMm / 2.0; 
        double slotTopYmm = slotBottomYmm + 86.0;
        double bossRectangleHeightMm = 160.0;
        double bossRectangleWidthMm = 8.0;
        double bossOuterDiameterOffsetMm = 9.0;
        double bossCenterlineAngleDeg = 2.96;
        double bossExtrusionDepthMm = 10.0;

        // Derived dimensions (m)
        double outerRadius = Mm(outerDiameterMm / 2.0);
        double innerRadius = Mm(innerDiameterMm / 2.0);
        double plateThickness = Mm(plateThicknessMm);
        double halfWidth = Mm(slotWidthMm / 2.0);
        double bottomY = Mm(slotBottomYmm);
        double topY = Mm(slotTopYmm);
        double centerY = (topY + bottomY) / 2.0;
        double leftX = -halfWidth;
        double bossRectangleHeight = Mm(bossRectangleHeightMm);
        double bossRectangleWidth = Mm(bossRectangleWidthMm);
        double bossOuterDiameterOffset = Mm(bossOuterDiameterOffsetMm);
        double bossCenterlineAngleRadians = bossCenterlineAngleDeg * Math.PI / 180.0;
        double bossExtrusionDepth = Mm(bossExtrusionDepthMm);
        double bossHalfWidth = bossRectangleWidth / 2.0;
        double bossTopCenterRadius = outerRadius - bossOuterDiameterOffset;
        double bossBottomCenterRadius = bossTopCenterRadius - bossRectangleHeight;
        double bossDirectionX = -Math.Sin(bossCenterlineAngleRadians);
        double bossDirectionY = Math.Cos(bossCenterlineAngleRadians);
        double bossNormalX = Math.Cos(bossCenterlineAngleRadians);
        double bossNormalY = Math.Sin(bossCenterlineAngleRadians);

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
            
            

            // Apply material to the  part
            PartDoc statorPart = swModel as PartDoc;
            statorPart.SetMaterialPropertyName2("", "", Name: "AISI 1020");

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

            // Create center rectangle for slot profile
            swSketchManager.CreateCenterRectangle(0, centerY, 0, halfWidth, topY, 0);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom edge");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(leftX + Mm(10), bottomY - Mm(10), 0);
            if (displayDim == null) throw new Exception("Could not create slot width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access slot width dimension");
            swDim.SystemValue = halfWidth * 2;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom-right point");
            selected = swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth, topY, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot top-right point");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(halfWidth + Mm(50), bottomY + Mm(50), 0);
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

            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot bottom edge for trimming");
            swSketchManager.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimClosest, 0, bottomY, 0);
            swModel.ClearSelection2(true);

            SketchSegment line3 = (SketchSegment)swSketchManager.CreateLine(-halfWidth, bottomY, 0, 0, 0, 0);
            SketchSegment line4 = (SketchSegment)swSketchManager.CreateLine(halfWidth, bottomY, 0, 0, 0, 0);
            if (line3 == null || line4 == null)
                throw new Exception("Could not create distance-sheet corner lines");
            line3.Select4(false, null);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

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
            if (myPattern == null)
                throw new Exception("Failed to create distance-sheet circular pattern");

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "FACE", (outerRadius + innerRadius) / 2.0, 0, plateThickness, false, 0, null, 0);
            if (!selected) throw new Exception("Could not reselect top face for boss sketch");

            bool bossSketchInferenceWasEnabled = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            swSketchManager.InsertSketch(true);

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
            SketchLine bossTopEdgeLine = bossTopEdge as SketchLine;
            if (bossTopEdgeLine == null)
                throw new Exception("Could not access boss top edge geometry");
            SketchPoint bossTopRightPoint = bossTopEdgeLine.GetEndPoint2();
            if (bossTopRightPoint == null)
                throw new Exception("Could not access boss top-right point");

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

            if (!bossRightEdge.Select4(false, null))
                throw new Exception("Could not select boss right edge for angle dimension");
            selected = swModel.Extension.SelectByID2("Y-Achse", "AXIS", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select Y axis for boss angle dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(bossTopCenterX + Mm(40), bossTopCenterY - Mm(40), 0);
            if (displayDim == null) throw new Exception("Could not create boss rectangle angle dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss rectangle angle dimension");
            swDim.SystemValue = bossCenterlineAngleRadians;
            swModel.ClearSelection2(true);

            if (!bossTopEdge.Select4(false, null))
                throw new Exception("Could not select boss top edge for width dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(
                bossTopCenterX + bossNormalX * Mm(25),
                bossTopCenterY + bossNormalY * Mm(25),
                0);
            if (displayDim == null) throw new Exception("Could not create boss width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss width dimension");
            swDim.SystemValue = bossRectangleWidth;
            swModel.ClearSelection2(true);

            double bossRightMidX = (topRightX + bottomRightX) / 2.0;
            double bossRightMidY = (topRightY + bottomRightY) / 2.0;
            if (!bossRightEdge.Select4(false, null))
                throw new Exception("Could not select boss right edge for height dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(
                bossRightMidX + bossNormalX * Mm(25),
                bossRightMidY + bossNormalY * Mm(25),
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
                bossTopCenterX / 2.0 + bossNormalX * Mm(20),
                bossTopCenterY / 2.0 + bossNormalY * Mm(20),
                0);
            if (displayDim == null) throw new Exception("Could not create boss top offset dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss top offset dimension");
            // Position the top edge at outer-radius minus the requested clearance.
            swDim.SystemValue = bossTopCenterRadius;
            swModel.ClearSelection2(true);

            if (!bossTopRightPoint.Select4(false, null))
                throw new Exception("Could not select boss top-right point for horizontal locating dimension");
            selected = swModel.Extension.SelectByID2("Y-Achse", "AXIS", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select Y axis for boss horizontal locating dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(topRightX / 2.0, topRightY + Mm(20), 0);
            if (displayDim == null) throw new Exception("Could not create boss horizontal locating dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access boss horizontal locating dimension");
            swDim.SystemValue = Math.Abs(topRightX);
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

            string savedPath;
            if (SaveToPdm)
            {
                savedPath = _pdm.SaveAsPdm(swModel, outFolder);
                Console.WriteLine($"Part saved to PDM: {savedPath}");
            }
            else
            {
                savedPath = Path.Combine(outFolder, "StatorDistanceBleche.SLDPRT");
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
