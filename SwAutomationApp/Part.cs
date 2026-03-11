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
        double bossCutOuterTabWidthMm = 2.0;
        double bossCutOuterTabHeightMm = 2.5;
        double bossCutTopShelfThicknessMm = 1.5;
        double bossCutInnerLegWidthMm = 1.5;
        double bossCutBoundaryExtensionMm = 2.0;
        int bossCircularPatternCount = 60;

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
        double bossCutOuterTabWidth = Mm(bossCutOuterTabWidthMm);
        double bossCutOuterTabHeight = Mm(bossCutOuterTabHeightMm);
        double bossCutTopShelfThickness = Mm(bossCutTopShelfThicknessMm);
        double bossCutInnerLegWidth = Mm(bossCutInnerLegWidthMm);
        double bossCutBoundaryExtension = Mm(bossCutBoundaryExtensionMm);

        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        try
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
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, topY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot top edge");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(leftX + Mm(10), topY + Mm(10), 0);
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

            SketchPoint line3OriginPoint = Math.Abs(line3Start.X) < Mm(0.001) && Math.Abs(line3Start.Y) < Mm(0.001) ? line3Start : line3End;
            SketchPoint line4OriginPoint = Math.Abs(line4Start.X) < Mm(0.001) && Math.Abs(line4Start.Y) < Mm(0.001) ? line4Start : line4End;
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

            if (!bossRightEdge.Select4(false, null))
                throw new Exception("Could not select boss right edge for height dimension");
            displayDim = (DisplayDimension)swModel.AddDimension2(
                ((topRightX + bottomRightX) / 2.0) + bossNormalX * Mm(25),
                ((topRightY + bottomRightY) / 2.0) + bossNormalY * Mm(25),
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

            object[] bossFeatureFaces = bossFeature.GetFaces() as object[];
            Face2 bossFrontFace = null;
            double bossFaceTolerance = Mm(0.01);
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
                ((cutP4[1] + cutP5[1]) / 2.0) - Mm(10),
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
                ((cutP2[1] + cutP3[1]) / 2.0) - Mm(10),
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
                ((cutP3[0] + cutP4[0]) / 2.0) + Mm(10),
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
                ((cutP5[0] + cutP6[0]) / 2.0) + Mm(10),
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
                ((cutP6[1] + cutP1[1]) / 2.0) + Mm(10),
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
                ((cutP1[0] + cutP2[0]) / 2.0) - Mm(10),
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

    public string Create_stator_end_sheet(string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
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

        // Derived dimensions (m)
        double outerRadius = Mm(outerDiameterMm / 2.0);
        double innerRadius = Mm(innerDiameterMm / 2.0);
        double plateThickness = Mm(plateThicknessMm);
        double halfWidth = Mm(slotWidthMm / 2.0);
        double bottomY = Mm(slotBottomYmm);
        double topY = Mm(slotTopYmm);
        double centerY = (topY + bottomY) / 2.0;
        double leftX = -halfWidth;

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
            swSketchManager.CreateCenterRectangle(0, centerY, 0, halfWidth, topY, 0);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, topY, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select slot top edge");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(leftX + Mm(10), topY + Mm(10), 0);
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

            SketchPoint line3OriginPoint = Math.Abs(line3Start.X) < Mm(0.001) && Math.Abs(line3Start.Y) < Mm(0.001) ? line3Start : line3End;
            SketchPoint line4OriginPoint = Math.Abs(line4Start.X) < Mm(0.001) && Math.Abs(line4Start.Y) < Mm(0.001) ? line4Start : line4End;
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
                throw new Exception("Failed to create circular slot pattern");

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
                _swApp.CloseDoc(swModel.GetTitle());
                Console.WriteLine("Part closed after creating.");
            }

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
            Console.WriteLine("Done!");

            return Path.GetFileName(savedPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fatal error: " + ex);
            try { _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true); } catch { }
            return null;
        }
    }

    public string Create_torsion_bar(string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
    {
        double Mm(double mm) => mm * MmToMeters;
        using var automationUi = BeginAutomationUiSuppression();
        bool SelectSketchByIndex(ModelDoc2 model, int index)
        {
            return model.Extension.SelectByID2($"Skizze{index}", "SKETCH", 0, 0, 0, false, 0, null, 0)
                || model.Extension.SelectByID2($"Sketch{index}", "SKETCH", 0, 0, 0, false, 0, null, 0);
        }

        // Main dimensions (mm) - change these only.
        double barLengthMm = 1074.0;
        double barHeightMm = 40.0;
        double barThicknessMm = 30.0;
        double holeCenterlineOffsetFromBottomMm = 20.0;
        double outerHoleEndOffsetMm = 30.0;
        double holePairSpacingMm = 315.0;
        double outerHoleDiameterMm = 10.0;
        double innerHoleDiameterMm = 16.0;
        double centerHoleDiameterMm = 16.0;
        string outerTapSizePrimary = "M10x1.5";
        string outerTapSizeFallback = "M10";
        string innerTapSizePrimary = "M16x2";
        string innerTapSizeFallback = "M16";

        string p0001ConfigName = "P0001";
        string p0002ConfigName = "P0002";

        // Derived dimensions (m)
        double barLength = Mm(barLengthMm);
        double barHeight = Mm(barHeightMm);
        double barThickness = Mm(barThicknessMm);
        double halfLength = barLength / 2.0;
        double halfHeight = barHeight / 2.0;
        double halfThickness = barThickness / 2.0;
        double holeCenterlineY = -halfHeight + Mm(holeCenterlineOffsetFromBottomMm);
        double outerHoleCenterX = halfLength - Mm(outerHoleEndOffsetMm);
        double innerHoleCenterX = outerHoleCenterX - Mm(holePairSpacingMm);
        double centerHoleRadius = Mm(centerHoleDiameterMm / 2.0);

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

            PartDoc torsionBarPart = swModel as PartDoc;
            torsionBarPart.SetMaterialPropertyName2("", "", Name: "AISI 1020");

            ConfigurationManager configMgr = swModel.ConfigurationManager;
            if (configMgr == null)
                throw new Exception("Could not access configuration manager");

            Configuration p0001Config = configMgr.ActiveConfiguration;
            if (p0001Config == null)
                throw new Exception("Could not access the active torsion-bar configuration");
            p0001Config.Name = p0001ConfigName;

            bool selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select Front Plane for torsion bar");

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCenterRectangle(0, 0, 0, halfLength, halfHeight, 0);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, halfHeight, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select torsion-bar top edge");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(0, halfHeight + Mm(20), 0);
            if (displayDim == null) throw new Exception("Could not create torsion-bar length dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access torsion-bar length dimension");
            swDim.SystemValue = barLength;

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfLength, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select torsion-bar right edge");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(halfLength + Mm(20), 0, 0);
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

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "FACE", 0, 0, halfThickness, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select torsion-bar front face for center-hole sketch");

            swSketchManager.InsertSketch(true);
            object centerHole = swSketchManager.CreateCircleByRadius(0, holeCenterlineY, 0, centerHoleRadius);
            if (centerHole == null)
                throw new Exception("Could not create torsion-bar center-hole sketch");

            swSketchManager.InsertSketch(true);
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
                Mm(outerHoleDiameterMm),
                -outerHoleCenterX,
                "left outer M10");
            Feature rightOuterTappedHoleFeature = CreateTappedHoleFeature(
                new[] { outerTapSizePrimary, outerTapSizeFallback },
                Mm(outerHoleDiameterMm),
                outerHoleCenterX,
                "right outer M10");
            Feature leftInnerTappedHoleFeature = CreateTappedHoleFeature(
                new[] { innerTapSizePrimary, innerTapSizeFallback },
                Mm(innerHoleDiameterMm),
                -innerHoleCenterX,
                "left inner M16");
            Feature rightInnerTappedHoleFeature = CreateTappedHoleFeature(
                new[] { innerTapSizePrimary, innerTapSizeFallback },
                Mm(innerHoleDiameterMm),
                innerHoleCenterX,
                "right inner M16");

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

            string savedPath;
            if (SaveToPdm)
            {
                savedPath = _pdm.SaveAsPdm(swModel, outFolder);
                Console.WriteLine($"Part saved to PDM: {savedPath}");
            }
            else
            {
                savedPath = Path.Combine(outFolder, "TorsionBar.SLDPRT");
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
        catch (Exception ex)
        {
            Console.WriteLine("Fatal error: " + ex);
            return null;
        }
    }

    public string Create_press_plate(string outFolder, bool closeAfterCreate = false, bool SaveToPdm = false)
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
        double ringInnerDiameterMm = 840.0;
        double plateInnerReferenceDiameterMm = 660.0;
        double ringThicknessMm = 2.0;
        double plateBodyThicknessMm = 10.0;
        double plateWidthMm = 6.0;
        int plateCount = 60;

        // Derived dimensions (m)
        double outerRadius = Mm(outerDiameterMm / 2.0);
        double ringInnerRadius = Mm(ringInnerDiameterMm / 2.0);
        double plateInnerRadius = Mm(plateInnerReferenceDiameterMm / 2.0);
        double ringThickness = Mm(ringThicknessMm);
        double plateBodyThickness = Mm(plateBodyThicknessMm);
        double plateHalfWidth = Mm(plateWidthMm / 2.0);
        double plateOuterY = outerRadius;
        double plateInnerY = plateInnerRadius;
        double plateCenterY = (plateOuterY + plateInnerY) / 2.0;
        double plateHeight = plateOuterY - plateInnerY;

        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;

        try
        {
            Dimension swDim = null;
            DisplayDimension displayDim = null;
            bool pressPlateSketchInferenceWasEnabled = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference);

            if (!Directory.Exists(outFolder))
                Directory.CreateDirectory(outFolder);

            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            if (swModel == null)
                throw new Exception("Failed to create new part");

            swSketchManager = swModel.SketchManager;

            PartDoc pressPlatePart = swModel as PartDoc;
            pressPlatePart.SetMaterialPropertyName2("", "", Name: "AISI 1020");

            bool selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select Front Plane for press plate");

            swSketchManager.InsertSketch(true);
            swSketchManager.CreateCircleByRadius(0, 0, 0, outerRadius);
            swSketchManager.CreateCircleByRadius(0, 0, 0, ringInnerRadius);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", outerRadius, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select press-plate outer circle");
            displayDim = (DisplayDimension)swModel.AddDimension2(outerRadius + Mm(20), Mm(20), 0);
            if (displayDim == null) throw new Exception("Could not create press-plate outer diameter dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate outer diameter dimension");
            swDim.SystemValue = Mm(outerDiameterMm);

            swModel.ClearSelection2(true);
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", ringInnerRadius, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select press-plate ring inner circle");
            displayDim = (DisplayDimension)swModel.AddDimension2(ringInnerRadius + Mm(20), Mm(20), 0);
            if (displayDim == null) throw new Exception("Could not create press-plate ring inner diameter dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate ring inner diameter dimension");
            swDim.SystemValue = Mm(ringInnerDiameterMm);

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

            SketchPoint LineStart(SketchLine line) => line.GetStartPoint2();
            SketchPoint LineEnd(SketchLine line) => line.GetEndPoint2();

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

            swModel.ClearSelection2(true);
            if (!(LineStart(topLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate body top-left point");
            if (!(LineEnd(topLine)?.Select4(true, null) ?? false))
                throw new Exception("Could not select press-plate body top-right point for width");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(Mm(10), plateOuterY + Mm(10), 0);
            if (displayDim == null) throw new Exception("Could not create press-plate body width dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate body width dimension");
            swDim.SystemValue = plateHalfWidth * 2.0;

            swModel.ClearSelection2(true);
            if (!(LineEnd(rightLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate body bottom-right point");
            if (!(LineStart(rightLine)?.Select4(true, null) ?? false))
                throw new Exception("Could not select press-plate body top-right point");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(plateHalfWidth + Mm(20), plateCenterY, 0);
            if (displayDim == null) throw new Exception("Could not create press-plate body height dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate body height dimension");
            swDim.SystemValue = plateHeight;

            swModel.ClearSelection2(true);
            if (!(LineEnd(topLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not select press-plate body top-right point for horizontal location");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for press-plate body horizontal location");
            displayDim = (DisplayDimension)swModel.AddHorizontalDimension2(plateHalfWidth / 2.0, plateOuterY + Mm(20), 0);
            if (displayDim == null) throw new Exception("Could not create press-plate body horizontal location dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate body horizontal location dimension");
            swDim.SystemValue = plateHalfWidth;
            swModel.ClearSelection2(true);

            if (!(LineEnd(topLine)?.Select4(false, null) ?? false))
                throw new Exception("Could not reselect press-plate body top-right point");
            selected = swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            if (!selected) throw new Exception("Could not select sketch origin for press-plate body top offset");
            displayDim = (DisplayDimension)swModel.AddVerticalDimension2(plateHalfWidth + Mm(20), plateOuterY / 2.0, 0);
            if (displayDim == null) throw new Exception("Could not create press-plate body top offset dimension");
            swDim = displayDim.GetDimension();
            if (swDim == null) throw new Exception("Could not access press-plate body top offset dimension");
            swDim.SystemValue = outerRadius;
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

            string savedPath;
            if (SaveToPdm)
            {
                savedPath = _pdm.SaveAsPdm(swModel, outFolder);
                Console.WriteLine($"Part saved to PDM: {savedPath}");
            }
            else
            {
                savedPath = Path.Combine(outFolder, "PressPlate.SLDPRT");
                swModel.SaveAs3(savedPath, 0, 1);
                Console.WriteLine($"Part saved locally: {savedPath}");
            }

            if (closeAfterCreate)
            {
                _swApp.CloseDoc(swModel.GetTitle());
                Console.WriteLine("Part closed after creating.");
            }

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, pressPlateSketchInferenceWasEnabled);
            Console.WriteLine("Done!");
            return Path.GetFileName(savedPath);
        }
        catch (Exception ex)
        {
            try { _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true); } catch { }
            Console.WriteLine("Fatal error: " + ex);
            return null;
        }
    }
}
