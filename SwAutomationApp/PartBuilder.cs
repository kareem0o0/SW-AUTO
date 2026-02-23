// Import base .NET types; alternative: remove if unused to speed build slightly.
using System;
// Import file system helpers for folders and paths.
using System.IO;
// Import SOLIDWORKS main COM interop types.
using SolidWorks.Interop.sldworks;
// Import SOLIDWORKS constant enums used for API calls.
using SolidWorks.Interop.swconst;

namespace SwAutomation;

// Contains all part-creation logic and part-plane helper functions.
public sealed class PartBuilder
{
    // SOLIDWORKS app instance used by all part operations.
    private readonly SldWorks _swApp;

    // Inject SOLIDWORKS session dependency.
    public PartBuilder(SldWorks swApp)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
    }

    // Generate a part with rectangular base and optional hole pattern.
    // This method explicitly describes the produced geometry.
    public string GenerateRectangularPartWithHoles(PartParameters parameters, string folder)
    {
        try
        {
            // Validate primary input values.
            if (string.IsNullOrWhiteSpace(parameters.Name) || parameters.WidthMm <= 0 || parameters.DepthMm <= 0 || parameters.HeightMm <= 0)
            {
                Console.WriteLine("Part generation skipped: invalid primary dimensions or name.");
                return string.Empty;
            }

            // Get the default Part template from SOLIDWORKS settings.
            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            // Create the new part document.
            ModelDoc2 swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            // Select the requested base plane and start a sketch.
            string[] preferredPlaneNames = GetSketchPlaneCandidates(parameters.SketchPlane);
            bool basePlaneSelected = SelectPartPlane(swModel, preferredPlaneNames, false);
            if (!basePlaneSelected)
            {
                // Fallback for unknown/localized plane names: select the first reference plane.
                basePlaneSelected = SelectFirstReferencePlane(swModel, false);
                if (!basePlaneSelected)
                {
                    Console.WriteLine("Part base sketch failed: requested base plane could not be selected.");
                    _swApp.CloseDoc(swModel.GetTitle());
                    return string.Empty;
                }
            }

            Console.WriteLine("Part base sketch plane selected successfully.");
            swModel.SketchManager.InsertSketch(true);

            // Draw the main rectangle (dimensions converted to meters); CENTERED.
            double x = parameters.WidthMm;
            double y = parameters.DepthMm;
            double z = parameters.HeightMm;
            swModel.SketchManager.CreateCenterRectangle(0, 0, 0, x / 2000, y / 2000, 0);

            // Disable "Input Dimension Value" to prevent popup dialogs.
            bool oldInputDimVal = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // Add Smart Dimensions to fully define the rectangle sketch.
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, (y / 2000), 0, false, 0, null, 0);
            swModel.AddDimension2(0, (y / 2000) + 0.01, 0);

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", (x / 2000), 0, 0, false, 0, null, 0);
            swModel.AddDimension2((x / 2000) + 0.01, 0, 0);

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            // Disable dimension popup for hole operations.
            bool oldInputDimValHoles = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // Create horizontal centerline to keep hole centers fully constrained.
            swModel.SketchManager.CreateLine(-(x / 2000), 0, 0, (x / 2000), 0, 0);
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", (x / 4000), 0, 0, false, 0, null, 0);
            swModel.SketchManager.CreateConstructionGeometry();

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", (x / 4000), 0, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");

            // Create hole circles if parameters request at least one valid hole.
            if (parameters.HoleDiameterMm > 0 && parameters.HoleCount > 0)
            {
                double spacingX = (x / 1000) / (parameters.HoleCount + 1);
                double startX = -(x / 2000);

                for (int i = 1; i <= parameters.HoleCount; i++)
                {
                    double centerX = startX + (spacingX * i);
                    double holeRadius = parameters.HoleDiameterMm / 2000;

                    swModel.SketchManager.CreateCircleByRadius(centerX, 0, 0, holeRadius);

                    swModel.Extension.SelectByID2("", "SKETCHSEGMENT", centerX + holeRadius, 0, 0, false, 0, null, 0);
                    swModel.AddDimension2(centerX + holeRadius + 0.01, 0.01, 0);

                    swModel.Extension.SelectByID2("", "SKETCHPOINT", centerX, 0, 0, false, 0, null, 0);
                    swModel.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
                    swModel.AddDimension2(centerX, -0.01, 0);

                    swModel.ClearSelection2(true);
                    swModel.Extension.SelectByID2("", "SKETCHPOINT", centerX, 0, 0, false, 0, null, 0);
                    swModel.Extension.SelectByID2("", "SKETCHSEGMENT", (x / 4000), 0, 0, true, 0, null, 0);
                    swModel.SketchAddConstraints("sgCOINCIDENT");
                }
            }

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimValHoles);

            // Extrude the sketch to create a 3D block.
            swModel.FeatureManager.FeatureExtrusion2(true, false, false, (int)swEndConditions_e.swEndCondBlind, 0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

            // Add a fillet on one selected edge.
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "EDGE", (x / 2000), (y / 2000), 0, false, 1, null, 0);
            swModel.FeatureManager.FeatureFillet3((int)swFeatureFilletOptions_e.swFeatureFilletUniformRadius, 0.005, 0, 0, 0, 0, 0, null, null, null, null, null, null, null);

            // Save part file and return full path.
            string fullPath = Path.Combine(folder, parameters.Name + ".SLDPRT");
            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
            return fullPath;
        }
        catch (Exception ex)
        {
            // Bubble a critical generation failure to Program.Main.
            throw new InvalidOperationException("Rectangular part generation failed for '" + parameters.Name + "'.", ex);
        }
    }

    // Backward-compatible alias.
    // Prefer using GenerateRectangularPartWithHoles for clearer intent.
    public string GeneratePart(PartParameters parameters, string folder)
    {
        return GenerateRectangularPartWithHoles(parameters, folder);
    }

    // Start a new blank part and begin a sketch on the requested base plane.
    // Caller can add sketch geometry, then use ExtrudeExistingSketchAndSave.
    public ModelDoc2 GeneratePart(string partName, SketchPlaneName sketchPlane = SketchPlaneName.Front)
    {
        if (string.IsNullOrWhiteSpace(partName)) throw new ArgumentException("Part name is required.", nameof(partName));

        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        ModelDoc2 swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

        string[] preferredPlaneNames = GetSketchPlaneCandidates(sketchPlane);
        bool basePlaneSelected = SelectPartPlane(swModel, preferredPlaneNames, false);
        if (!basePlaneSelected)
        {
            basePlaneSelected = SelectFirstReferencePlane(swModel, false);
            if (!basePlaneSelected)
            {
                throw new InvalidOperationException("Failed to select base sketch plane while starting part.");
            }
        }

        swModel.SketchManager.InsertSketch(true);
        Console.WriteLine("Started part '" + partName + "' sketch on " + sketchPlane + " plane.");
        return swModel;
    }

    // Extrude an existing sketch in the provided part document and save.
    // Expected flow:
    // 1) GeneratePart(partName, plane)
    // 2) create sketch entities in the active sketch
    // 3) call this method to close sketch (if open), extrude, and save.
    public string ExtrudeExistingSketchAndSave(
        ModelDoc2 partModel,
        string partName,
        string outputFolder,
        double depthMm,
        bool midPlane = false)
    {
        if (partModel == null) throw new ArgumentNullException(nameof(partModel));
        if (string.IsNullOrWhiteSpace(partName)) throw new ArgumentException("Part name is required.", nameof(partName));
        if (string.IsNullOrWhiteSpace(outputFolder)) throw new ArgumentException("Output folder is required.", nameof(outputFolder));
        if (depthMm <= 0) throw new ArgumentOutOfRangeException(nameof(depthMm), "depthMm must be > 0.");

        if (!Directory.Exists(outputFolder)) Directory.CreateDirectory(outputFolder);

        // If sketch is currently open, close it before extrusion.
        Sketch activeSketch = partModel.GetActiveSketch2();
        if (activeSketch != null)
        {
            partModel.SketchManager.InsertSketch(true);
        }

        Feature sketchFeature = GetLastSketchFeature(partModel);
        if (sketchFeature == null || !sketchFeature.Select2(false, 0))
        {
            throw new InvalidOperationException("No valid existing sketch was found for extrusion.");
        }

        int endCondition = midPlane
            ? (int)swEndConditions_e.swEndCondMidPlane
            : (int)swEndConditions_e.swEndCondBlind;

        Feature createdExtrude = partModel.FeatureManager.FeatureExtrusion2(
            true, false, false,
            endCondition, 0,
            depthMm / 1000.0, 0,
            false, false, false, false,
            0, 0, false, false, false, false,
            true, true, true, 0, 0, false);

        if (createdExtrude == null)
        {
            throw new InvalidOperationException("Extrusion failed for the existing sketch.");
        }

        partModel.ForceRebuild3(false);

        string fullPath = Path.Combine(outputFolder, partName + ".SLDPRT");
        partModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);

        return fullPath;
    }

    // Create a new part and add two offset planes on opposite sides of the Right Plane.
    public void CreatePartWithOffsetPlanes(double offsetDistance)
    {
        try
        {
            bool oldInputDimVal = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            Feature plane1 = swModel.CreatePlaneAtOffset3(offsetDistance, false, false);
            if (plane1 != null) plane1.Name = "Offset_Right";

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            Feature plane2 = swModel.CreatePlaneAtOffset3(offsetDistance, true, false);
            if (plane2 != null) plane2.Name = "Offset_Left";

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);
            swModel.ForceRebuild3(false);
        }
        catch (Exception ex)
        {
            // Raise critical errors to entry point.
            throw new InvalidOperationException("Failed to create part with offset planes.", ex);
        }
    }

    // Create a reference offset plane in an existing component.
    public void CreateReferenceOffsetPlane(string componentPath, string planeName, double offsetDistance)
    {
        try
        {
            bool oldInputDimVal = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            int errors = 0;
            int warnings = 0;
            ModelDoc2 swModel = _swApp.OpenDoc6(componentPath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            if (swModel == null)
            {
                _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);
                Console.WriteLine("CreateReferenceOffsetPlane skipped: target part could not be opened.");
                return;
            }

            string standardPlaneName = planeName.Contains("Plane", StringComparison.OrdinalIgnoreCase) ? planeName : planeName + " Plane";
            bool status = swModel.Extension.SelectByID2(standardPlaneName, "PLANE", 0, 0, 0, false, 0, null, 0);
            if (status)
            {
                Feature newPlane = swModel.CreatePlaneAtOffset3(offsetDistance, false, false);
                if (newPlane != null) newPlane.Name = planeName + "_Offset";
            }
            else
            {
                Console.WriteLine("CreateReferenceOffsetPlane warning: base plane could not be selected.");
            }

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);
            swModel.ForceRebuild3(false);
            swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create reference offset plane for '" + componentPath + "'.", ex);
        }
    }

    // Create a rectangular part centered at origin with fully defined sketch and mid-plane extrusion.
    public void CreateCenteredRectangularPart(string partName, double x, double y, double z, string outputFolder)
    {
        try
        {
            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            swModel.SketchManager.CreateCenterRectangle(0, 0, 0, x / 2000, y / 2000, 0);

            bool oldInputDimVal = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, 0, (y / 2000), false, 0, null, 0);
            swModel.AddDimension2(0, 0, (y / 2000) + 0.02);

            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", (x / 2000), 0, 0, false, 0, null, 0);
            swModel.AddDimension2((x / 2000) + 0.02, 0, 0);

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            swModel.FeatureManager.FeatureExtrusion2(true, false, false, (int)swEndConditions_e.swEndCondMidPlane, 0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

            string fullPath = Path.Combine(outputFolder, partName + ".SLDPRT");
            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create centered rectangular part '" + partName + "'.", ex);
        }
    }

    // Create a circular part centered at origin with fully defined sketch and mid-plane extrusion.
    public void CreateCenteredCircularPart(string partName, double diameter, double z, string outputFolder)
    {
        try
        {
            string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);

            swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            swModel.SketchManager.CreateCircleByRadius(0, 0, 0, diameter / 2000);

            bool oldInputDimVal = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", diameter / 2000, 0, 0, false, 0, null, 0);
            swModel.AddDimension2((diameter / 2000) + 0.02, 0.02, 0);

            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            swModel.FeatureManager.FeatureExtrusion2(true, false, false, (int)swEndConditions_e.swEndCondMidPlane, 0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

            string fullPath = Path.Combine(outputFolder, partName + ".SLDPRT");
            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create centered circular part '" + partName + "'.", ex);
        }
    }

    // Mirror a selected part feature/body about a selected mirror plane.
    // Example:
    //   partBuilder.MirrorPartFeature(
    //       partPath,
    //       new SketchEntityReference("PLANE", 0, 0, 0, "Front Plane"),
    //       new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"));
    public void MirrorPartFeature(
        string partPath,
        SketchEntityReference mirrorPlaneReference,
        SketchEntityReference seedFeatureReference,
        bool geometryPattern = true,
        bool merge = true,
        bool knit = false,
        int scopeOptions = 0)
    {
        try
        {
            ModelDoc2 swModel = OpenPartDocument(partPath);
            if (swModel == null)
            {
                Console.WriteLine("MirrorPartFeature skipped: target part could not be opened.");
                return;
            }

            swModel.ClearSelection2(true);

            bool seedSelected = SelectInPart(swModel, seedFeatureReference, false, 0);
            if (!seedSelected)
            {
                Console.WriteLine("MirrorPartFeature skipped: failed to select seed feature/body.");
                return;
            }

            bool planeSelected = SelectInPart(swModel, mirrorPlaneReference, true, 0);
            if (!planeSelected)
            {
                Console.WriteLine("MirrorPartFeature skipped: failed to select mirror plane.");
                return;
            }

            Feature mirrored = swModel.FeatureManager.InsertMirrorFeature2(false, geometryPattern, merge, knit, scopeOptions);
            if (mirrored == null)
            {
                Console.WriteLine("MirrorPartFeature warning: mirror feature could not be created.");
            }

            SaveAndRebuildPart(swModel);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create mirrored part feature for '" + partPath + "'.", ex);
        }
    }

    // Create linear pattern for a selected part feature/body.
    // Example:
    //   partBuilder.CreatePartLinearPattern(
    //       partPath,
    //       new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"),
    //       new SketchEntityReference("EDGE", 0, 0, 0),
    //       4,
    //       20);
    public void CreatePartLinearPattern(
        string partPath,
        SketchEntityReference seedFeatureReference,
        SketchEntityReference direction1Reference,
        int count1,
        double spacing1Mm,
        SketchEntityReference direction2Reference = default,
        int count2 = 1,
        double spacing2Mm = 0,
        bool flipDir1 = false,
        bool flipDir2 = false)
    {
        try
        {
            if (count1 < 1) throw new ArgumentOutOfRangeException(nameof(count1), "count1 must be >= 1.");
            if (count2 < 1) throw new ArgumentOutOfRangeException(nameof(count2), "count2 must be >= 1.");
            if (spacing1Mm < 0) throw new ArgumentOutOfRangeException(nameof(spacing1Mm), "spacing1Mm must be >= 0.");
            if (spacing2Mm < 0) throw new ArgumentOutOfRangeException(nameof(spacing2Mm), "spacing2Mm must be >= 0.");

            ModelDoc2 swModel = OpenPartDocument(partPath);
            if (swModel == null)
            {
                Console.WriteLine("CreatePartLinearPattern skipped: target part could not be opened.");
                return;
            }

            swModel.ClearSelection2(true);

            bool seedSelected = SelectInPart(swModel, seedFeatureReference, false, 0);
            if (!seedSelected)
            {
                Console.WriteLine("CreatePartLinearPattern skipped: failed to select seed feature/body.");
                return;
            }

            bool dir1Selected = SelectInPart(swModel, direction1Reference, true, 0);
            if (!dir1Selected)
            {
                Console.WriteLine("CreatePartLinearPattern skipped: failed to select first direction reference.");
                return;
            }

            if (direction2Reference.IsValid())
            {
                bool dir2Selected = SelectInPart(swModel, direction2Reference, true, 0);
                if (!dir2Selected)
                {
                    Console.WriteLine("CreatePartLinearPattern skipped: failed to select second direction reference.");
                    return;
                }
            }

            Feature createdPattern = swModel.FeatureManager.FeatureLinearPattern(
                count1,
                spacing1Mm / 1000.0,
                count2,
                spacing2Mm / 1000.0,
                flipDir1,
                flipDir2,
                string.Empty,
                string.Empty);

            if (createdPattern == null)
            {
                Console.WriteLine("CreatePartLinearPattern warning: linear pattern feature could not be created.");
            }

            SaveAndRebuildPart(swModel);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create linear pattern for '" + partPath + "'.", ex);
        }
    }

    // Create circular pattern for a selected part feature/body.
    // Example:
    //   partBuilder.CreatePartCircularPattern(
    //       partPath,
    //       new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"),
    //       new SketchEntityReference("AXIS", 0, 0, 0),
    //       6,
    //       60);
    public void CreatePartCircularPattern(
        string partPath,
        SketchEntityReference seedFeatureReference,
        SketchEntityReference axisReference,
        int count,
        double spacingDeg,
        bool flipDir = false)
    {
        try
        {
            if (count < 1) throw new ArgumentOutOfRangeException(nameof(count), "count must be >= 1.");

            ModelDoc2 swModel = OpenPartDocument(partPath);
            if (swModel == null)
            {
                Console.WriteLine("CreatePartCircularPattern skipped: target part could not be opened.");
                return;
            }

            swModel.ClearSelection2(true);

            bool seedSelected = SelectInPart(swModel, seedFeatureReference, false, 0);
            if (!seedSelected)
            {
                Console.WriteLine("CreatePartCircularPattern skipped: failed to select seed feature/body.");
                return;
            }

            bool axisSelected = SelectInPart(swModel, axisReference, true, 0);
            if (!axisSelected)
            {
                Console.WriteLine("CreatePartCircularPattern skipped: failed to select axis reference.");
                return;
            }

            Feature createdPattern = swModel.FeatureManager.FeatureCircularPattern(
                count,
                spacingDeg * Math.PI / 180.0,
                flipDir,
                string.Empty);

            if (createdPattern == null)
            {
                Console.WriteLine("CreatePartCircularPattern warning: circular pattern feature could not be created.");
            }

            SaveAndRebuildPart(swModel);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to create circular pattern for '" + partPath + "'.", ex);
        }
    }

    // Try selecting a part plane from multiple localized names.
    private static bool SelectPartPlane(ModelDoc2 partModel, string[] planeNames, bool append)
    {
        for (int i = 0; i < planeNames.Length; i++)
        {
            bool selected = partModel.Extension.SelectByID2(planeNames[i], "PLANE", 0, 0, 0, append, 0, null, 0);
            if (selected) return true;
        }

        return false;
    }

    // Return the latest sketch feature in the feature tree.
    private static Feature GetLastSketchFeature(ModelDoc2 model)
    {
        Feature feat = model.FirstFeature();
        Feature lastSketch = null;
        while (feat != null)
        {
            string type = feat.GetTypeName2();
            if (string.Equals(type, "ProfileFeature", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "3DProfileFeature", StringComparison.OrdinalIgnoreCase))
            {
                lastSketch = feat;
            }

            feat = feat.GetNextFeature();
        }

        return lastSketch;
    }

    // Open an existing part document silently.
    private ModelDoc2 OpenPartDocument(string partPath)
    {
        if (string.IsNullOrWhiteSpace(partPath))
        {
            throw new ArgumentException("Part path is required.", nameof(partPath));
        }

        int errors = 0;
        int warnings = 0;
        return _swApp.OpenDoc6(
            partPath,
            (int)swDocumentTypes_e.swDocPART,
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
            string.Empty,
            ref errors,
            ref warnings);
    }

    // Select an entity in part using typed reference data.
    private static bool SelectInPart(ModelDoc2 model, SketchEntityReference reference, bool append, int mark)
    {
        if (!reference.IsValid()) return false;
        return model.Extension.SelectByID2(
            reference.Name ?? string.Empty,
            reference.Type,
            reference.XMm / 1000.0,
            reference.YMm / 1000.0,
            reference.ZMm / 1000.0,
            append,
            mark,
            null,
            0);
    }

    // Rebuild and save part document silently.
    private static void SaveAndRebuildPart(ModelDoc2 model)
    {
        model.ForceRebuild3(false);
        int errors = 0;
        int warnings = 0;
        model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
    }

    // Return localized plane-name candidates for requested sketch plane option.
    private static string[] GetSketchPlaneCandidates(SketchPlaneName sketchPlane)
    {
        switch (sketchPlane)
        {
            case SketchPlaneName.Top:
                return new[] { "Top Plane", "Top", "Ebene Oben", "Oben" };
            case SketchPlaneName.Right:
            case SketchPlaneName.Side:
                return new[] { "Right Plane", "Right", "Ebene Rechts", "Rechts" };
            case SketchPlaneName.Front:
            default:
                return new[] { "Front Plane", "Front", "Ebene Vorne", "Vorne" };
        }
    }

    // Fallback helper: pick first reference plane in feature tree.
    private static bool SelectFirstReferencePlane(ModelDoc2 partModel, bool append)
    {
        Feature feat = partModel.FirstFeature();
        while (feat != null)
        {
            string featureType = feat.GetTypeName2();
            if (string.Equals(featureType, "RefPlane", StringComparison.OrdinalIgnoreCase))
            {
                bool selected = feat.Select2(append, 0);
                if (selected)
                {
                    Console.WriteLine("Part base sketch plane selected by fallback feature scan.");
                    return true;
                }
            }

            feat = feat.GetNextFeature();
        }

        return false;
    }
}
