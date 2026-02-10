// Import base .NET types; alternative: remove if unused to speed build slightly.
using System;
// Import file system helpers for folders and paths.
using System.IO;
// Import SOLIDWORKS main COM interop types.
using SolidWorks.Interop.sldworks;
// Import SOLIDWORKS constant enums used for API calls.
using SolidWorks.Interop.swconst;

// Removes the .NET 8 warning by specifying this code is Windows-only; alternative: set in project file.
[assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")]

// Namespace groups related automation code; alternative: match your org/company root namespace.
namespace SwAutomation
// Begin namespace scope.
{
    // Program entry type; alternative: use a separate App class for DI/testing.
    class Program
    // Begin class scope.
    {
        // App entry point; alternative: make async Task Main if you need awaits.
        // Args: args = command-line arguments passed to the app (unused here, but can hold paths/options).
        static void Main(string[] args)
        // Begin Main method scope.
        {
            string outFolder = @"D:\work\birr machines\parts";
            // Keep the essential SOLIDWORKS connection in place; uncomment one of the method calls below to run it.
            try
            {
                // Inform the user we're about to connect to SOLIDWORKS.
                Console.WriteLine("Connecting to SOLIDWORKS...");
                // Create or connect to SOLIDWORKS via COM ProgID.
                // Args: "SldWorks.Application"=COM ProgID to locate SOLIDWORKS; swApp=active SOLIDWORKS application instance.
                SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
                // Show the SOLIDWORKS UI; set false for headless/batch if supported.
                swApp.Visible = true;

                // Example: Generate two parts and an assembly (uncomment after creating swApp and output folder).
                // Args: outFolder=directory where files are saved.
                // if (!Directory.Exists(outFolder)) Directory.CreateDirectory(outFolder); // Ensure folder exists.
                // string partAPath = GeneratePart(swApp, "Part_A", "100", "50", "10", "5", "3", outFolder); // 100x50x10, 3 holes.
                // string partBPath = GeneratePart(swApp, "Part_B", "80", "80", "20", "4", "2", outFolder); // 80x80x20, 2 holes.
                // GenerateAssembly(swApp, partAPath, partBPath, outFolder); // Assemble and mate the two parts.

                // Example: Create offset planes in a new part (uncomment after creating swApp).
                // Args: offsetDistance=distance in meters (0.05 = 50mm).
                // CreatePartWithOffsetPlanes(swApp, 0.05);

                // --- NEW EXAMPLES ---

                // Example: Create a Reference Offset Plane in an existing part.
                // Inputs: Path to part, Plane Name ("Front", "Top", "Right"), Offset in meters.
                // CreateReferenceOffsetPlane(swApp, @"D:\work\birr machines\parts\Part_A.SLDPRT", "Front", 0.10);

                // Example: Create a Centered Rectangular Part.
                // Inputs: Filename, Width (mm), Depth (mm), Height (mm), Output Folder.
                 CreateCenteredRectangularPart(swApp, "RectPart_Centered", 120, 60, 20, outFolder);

                // Example: Create a Centered Circular Part.
                // Inputs: Filename, Diameter (mm), Height (mm), Output Folder.
                // CreateCenteredCircularPart(swApp, "CircPart_Centered", 50, 100, outFolder);

                // Note: For production use, add cleanup (e.g., swApp.ExitApp()) when done.
            }
            catch (Exception ex)
            {
                // Report any error during connection or method execution.
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        // Generate a part with rectangular base and hole pattern; alternative: accept numeric types not strings.
        // Args: swApp=SOLIDWORKS application; name=part base filename; sx=x size mm (string); sy=y size mm (string); sz=z size mm (string); sDia=hole diameter mm (string); sCount=hole count (string); folder=output directory.
        static string GeneratePart(SldWorks swApp, string name, string sx, string sy, string sz, string sDia, string sCount, string folder)
        // Begin GeneratePart method scope.
        {
            // Convert string inputs to doubles (meters = mm / 1000); alternative: use CultureInfo.InvariantCulture.
            // Args: sx/sy/sz = strings to parse; out x/y/z = parsed numeric results.
            if (!double.TryParse(sx, out double x) || !double.TryParse(sy, out double y) || !double.TryParse(sz, out double z)) return "";

            // Get the default Part template from SOLIDWORKS settings; alternative: use a custom template path.
            // Args: swDefaultTemplatePart = preference key for the default part template.
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            // Create the new document; alternative: specify document units in template.
            // Args: template=template path; 0=paper size (unused for part); 0=width (unused); 0=height (unused).
            ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);

            // Select the Front Plane and start a sketch; alternative: use Top Plane based on your modeling standard.
            // Args: "Front Plane"=entity name; "PLANE"=entity type; 0,0,0=selection point; false=do not append; 0=mark; null=callout; 0=select option.
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            // Enter sketch mode; alternative: check if already in sketch before inserting.
            // Args: true=toggle/enter sketch (SOLIDWORKS treats this as start/exit based on current state).
            swModel.SketchManager.InsertSketch(true);

            // Draw the main rectangle (Dimensions converted to meters); CENTERED.
            // Args: 0,0,0=center; x/2000,y/2000,0=corner offset (half width/height).
            // Note: Creating a center rectangle ensures symmetry directly.
            swModel.SketchManager.CreateCenterRectangle(0, 0, 0, x / 2000, y / 2000, 0);

            // Disable "Input Dimension Value" to prevent popup dialogs; automation must run without pauses.
            // Args: swInputDimValOnCreate=preference to toggle; false=disable.
            bool oldInputDimVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // Add Smart Dimensions to fully define the sketch.
            // 1. Dimension Width (Top Line).
            swModel.ClearSelection2(true); // Clear created rectangle entities.
            // Select the top horizontal line for width dimension.
            // Args: 0,(y/2000),0 = point on the top edge.
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, (y / 2000), 0, false, 0, null, 0);
            // Add dimension.
            // Args: 0,(y/2000)+0.01,0 = text position.
            swModel.AddDimension2(0, (y / 2000) + 0.01, 0);

            // 2. Dimension Height (Right Line).
            swModel.ClearSelection2(true);
            // Select the right vertical line for height dimension.
            // Args: (x/2000),0,0 = point on the right edge.
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", (x / 2000), 0, 0, false, 0, null, 0);
            // Add dimension.
            // Args: (x/2000)+0.01,0,0 = text position.
            swModel.AddDimension2((x / 2000) + 0.01, 0, 0);

            // Restore the original preference setting.
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            // Disable "Input Dimension Value" to prevent popup dialogs; automation must run without pauses.
            // Args: swInputDimValOnCreate=preference to toggle; false=disable.
            bool oldInputDimValHoles = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // Create Holes: Calculate spacing based on count and draw circles; alternative: use Hole Wizard API.
            // Args: sDia=string hole diameter; out hDia=parsed diameter; sCount=string count; out hCount=parsed count.
            if (double.TryParse(sDia, out double hDia) && int.TryParse(sCount, out int hCount) && hCount > 0)
            // Begin hole creation block.
            {
                // Compute equal spacing along X; starting from center.
                // Full width is x/1000.
                double spacingX = (x / 1000) / (hCount + 1);
                // Start X position (left side).
                double startX = -(x / 2000);

                // Loop over each hole.
                for (int i = 1; i <= hCount; i++)
                {
                    // Calculate center X for this hole.
                    double centerX = startX + (spacingX * i);

                    // Create circle.
                    swModel.SketchManager.CreateCircleByRadius(centerX, 0, 0, (hDia / 2000)); // centered on Y axis (y=0)

                    // Add Smart Dimension for Diameter.
                    // Select the circle arc.
                    // Args: centerX + (hDia/2000), 0, 0 = point on circumference.
                    swModel.Extension.SelectByID2("", "SKETCHSEGMENT", centerX + (hDia / 2000), 0, 0, false, 0, null, 0);
                    // Add dimension text.
                    swModel.AddDimension2(centerX + (hDia / 2000) + 0.01, 0.01, 0);

                    // Add Smart Dimension for X Position (from Origin).
                    // Select the circle center.
                    swModel.Extension.SelectByID2("", "SKETCHPOINT", centerX, 0, 0, false, 0, null, 0);
                    // Select the Origin.
                    swModel.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
                    // Add horizontal dimension.
                    swModel.AddDimension2(centerX, -0.01, 0);

                    // Add Smart Dimension for Y Position (Vertical from Origin) to lock it to 0.
                    // Select the circle center.
                    swModel.Extension.SelectByID2("", "SKETCHPOINT", centerX, 0, 0, false, 0, null, 0);
                    // Select the Origin.
                    swModel.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, true, 0, null, 0);
                    // Add vertical dimension; this will be 0, effectively locking it to the axis.
                    // Alternatively, could use a geometric relation (Coincident) to the X-axis/Top Plane normal.
                    swModel.AddDimension2(centerX + 0.01, 0, 0);
                }
            }

            // Restore the original preference setting.
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimValHoles);

            // Extrude the sketch to create a 3D block; alternative: use FeatureExtrusion3 for newer options.
            // Args (FeatureExtrusion2):
            // 1) true=solid feature, 2) false=do not thin, 3) false=not a cut,
            // 4) swEndCondBlind=end condition, 5) 0=dir2 end condition,
            // 6) z/1000=depth1 (m), 7) 0=depth2,
            // 8) false=draft outward, 9) false=draft outward dir2,
            // 10) false=add draft, 11) false=add draft dir2,
            // 12) 0=draft angle, 13) 0=draft angle dir2,
            // 14) false=offset start, 15) false=offset end,
            // 16) false=merge result?, 17) false=use auto select?,
            // 18) true=keep body, 19) true=auto select direction,
            // 20) true=feature scope auto, 21) 0=thin wall thickness, 22) 0=thin wall thickness2,
            // 23) false=thin wall reversed.
            swModel.FeatureManager.FeatureExtrusion2(true, false, false, (int)swEndConditions_e.swEndCondBlind, 0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

            // Save the part to the specified folder; alternative: check SaveAs3 return codes.
            // Args: folder=parent directory; name+".SLDPRT"=file name.
            string fullPath = Path.Combine(folder, name + ".SLDPRT");
            // Save silently to current version; alternative: use swSaveAsOptions_e.swSaveAsOptions_Copy.
            // Args: fullPath=target file path; swSaveAsCurrentVersion=save version; swSaveAsOptions_Silent=no UI dialogs.
            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);

            // Return the saved file path; alternative: return ModelDoc2 for further edits.
            return fullPath;
        }

        // Build an assembly that inserts two parts and mates planes; alternative: fully define with additional mates.
        // Args: swApp=SOLIDWORKS application; partAPath=full path to part A; partBPath=full path to part B; folder=output directory.
        static void GenerateAssembly(SldWorks swApp, string partAPath, string partBPath, string folder)
        // Begin GenerateAssembly method scope.
        {
            // Disable "Input Dimension Value" to avoid pauses during component insertion or mating interactions.
            bool oldInputDimVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // Get the default Assembly template and create a new assembly document; alternative: use a custom assembly template.
            // Args: swDefaultTemplateAssembly = preference key for the default assembly template.
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            // Create assembly document; alternative: set units or display states in template.
            // Args: template=template path; 0=paper size (unused for assembly); 0=width (unused); 0=height (unused).
            ModelDoc2 assyModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);
            // Cast to AssemblyDoc for assembly-specific APIs; alternative: verify cast success.
            AssemblyDoc swAssy = (AssemblyDoc)assyModel;

            // 1. INSERT COMPONENTS; alternative: add components at specific transforms.
            // AddComponent5 returns the Component2 object, which allows us to get the exact instance name (e.g., Part_A-1); alternative: use AddComponent4 for older versions.
            // Args: partAPath=component file; CurrentSelectedConfig=use active config; ""=config name; false=do not suppress; ""=component name; 0,0,0=insert at origin.
            Component2 compA = swAssy.AddComponent5(partAPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
            // Insert second component with a Z offset; alternative: place by transform matrix for precise positioning.
            // Args: partBPath=component file; CurrentSelectedConfig=use active config; ""=config name; false=do not suppress; ""=component name; 0,0,0.05=insert at Z offset 0.05m.
            Component2 compB = swAssy.AddComponent5(partBPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0.05);

            // Capture names for selection strings; alternative: use IComponent2::GetPathName for unique IDs.
            string assemblyName = assyModel.GetTitle();
            // Name of component instance A; alternative: use Name2 after checking for null.
            string partAName = compA.Name2;
            // Name of component instance B; alternative: store as Component2 reference only.
            string partBName = compB.Name2;

            // 2. MATE FRONT PLANES (Aligns them in the Z-direction); alternative: use coincident faces.
            // Args: true=clear all selection marks.
            assyModel.ClearSelection2(true);

            // Format: "PlaneName@InstanceName@AssemblyName"; alternative: select by ray with SelectByID2 on face.
            // Args: name=selection string; "PLANE"=entity type; 0,0,0=selection point; false=do not append; 1=mark; null=callout; 0=select option.
            bool status1 = assyModel.Extension.SelectByID2("Front Plane@" + partAName + "@" + assemblyName, "PLANE", 0, 0, 0, false, 1, null, 0);
            // Select the second plane, append to selection set; alternative: use SelectByID2 with mark values.
            // Args: name=selection string; "PLANE"=entity type; 0,0,0=selection point; true=append to selection; 1=mark; null=callout; 0=select option.
            bool status2 = assyModel.Extension.SelectByID2("Front Plane@" + partBName + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);

            // Only add mate if both selections succeeded; alternative: throw on failure for easier debugging.
            if (status1 && status2)
            // Begin mate creation block.
            {
                // Output parameter for mate creation result; alternative: inspect mateError codes for diagnostics.
                int mateError = 0;
                // Add coincident mate; alternative: use swMateDISTANCE with an offset.
                // Args: swMateCOINCIDENT=mate type; swMateAlignALIGNED=alignment; false=flip; 0..0=distances/angles; false=use for reference; out mateError=error code.
                swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out mateError);
            }

            // 3. MATE TOP PLANES (Aligns them in the Y-direction); alternative: mate edges or axes instead.
            // Args: true=clear all selection marks.
            assyModel.ClearSelection2(true);
            // Select top plane of component A; alternative: select reference geometry by name.
            // Args: name=selection string; "PLANE"=entity type; 0,0,0=selection point; false=do not append; 1=mark; null=callout; 0=select option.
            status1 = assyModel.Extension.SelectByID2("Top Plane@" + partAName + "@" + assemblyName, "PLANE", 0, 0, 0, false, 1, null, 0);
            // Select top plane of component B; alternative: pick face IDs for robustness.
            // Args: name=selection string; "PLANE"=entity type; 0,0,0=selection point; true=append; 1=mark; null=callout; 0=select option.
            status2 = assyModel.Extension.SelectByID2("Top Plane@" + partBName + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);

            // Only add mate if both selections succeeded; alternative: log detailed selection errors.
            if (status1 && status2)
            // Begin mate creation block.
            {
                // Output parameter for mate creation result; alternative: inspect and retry on errors.
                int mateError = 0;
                // Add coincident mate; alternative: use swMatePARALLEL for less constraint.
                // Args: swMateCOINCIDENT=mate type; swMateAlignALIGNED=alignment; false=flip; 0..0=distances/angles; false=use for reference; out mateError=error code.
                swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out mateError);
            }

            // Restore the original preference setting.
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            // Final Rebuild and Save; alternative: check rebuild errors and resolve dangling mates.
            // Args: false=full rebuild (not top-only).
            assyModel.ForceRebuild3(false);
            // Compose assembly path; alternative: include timestamp/version in filename.
            // Args: folder=parent directory; "Final_Assembly.SLDASM"=file name.
            string assemblyPath = Path.Combine(folder, "Final_Assembly.SLDASM");
            // Save silently; alternative: use Save3 and inspect errors/warnings.
            // Args: assemblyPath=target file path; swSaveAsCurrentVersion=save version; swSaveAsOptions_Silent=no UI dialogs.
            assyModel.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
        }

        // Create a new part and add two offset planes on opposite sides of the Right Plane.
        // Args: swApp=SOLIDWORKS application; offsetDistance=distance in meters for plane offset (e.g., 0.05 = 50mm).
        static void CreatePartWithOffsetPlanes(SldWorks swApp, double offsetDistance)
        // Begin CreatePartWithOffsetPlanes method scope.
        {
            // Disable "Input Dimension Value" to prevent popup dialogs; automation must run without pauses.
            bool oldInputDimVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // 1. Create a new part document.
            // Args: swDefaultTemplatePart = preference key for the default part template.
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            // Args: template=template path; 0=paper size (unused for part); 0=width (unused); 0=height (unused).
            ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);

            // 2. Create the first offset plane (standard direction).
            // Args: true=clear all selection marks.
            swModel.ClearSelection2(true);
            // Select the Right Plane as our base.
            // Args: "Right Plane"=entity name; "PLANE"=entity type; 0,0,0=selection point; false=do not append; 0=mark; null=callout; 0=select option.
            swModel.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, 0, null, 0);

            // CreatePlaneAtOffset3(Distance, ReverseDirection, Options).
            // Args: offsetDistance=offset in meters; false=standard direction; false=default options.
            Feature plane1 = swModel.CreatePlaneAtOffset3(offsetDistance, false, false);
            // If created, name it for easy selection later.
            if (plane1 != null) plane1.Name = "Offset_Right";

            // 3. Create the second offset plane (reversed direction).
            // Args: true=clear all selection marks.
            swModel.ClearSelection2(true);
            // Select the Right Plane again as the base.
            // Args: "Right Plane"=entity name; "PLANE"=entity type; 0,0,0=selection point; false=do not append; 0=mark; null=callout; 0=select option.
            swModel.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, 0, null, 0);

            // Use true to flip the direction to the other side.
            // Args: offsetDistance=offset in meters; true=reverse direction; false=default options.
            Feature plane2 = swModel.CreatePlaneAtOffset3(offsetDistance, true, false);
            // If created, name it for easy selection later.
            if (plane2 != null) plane2.Name = "Offset_Left";

            // Restore the original preference setting.
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            // Rebuild the model to ensure planes are fully created.
            // Args: false=full rebuild (not top-only).
            swModel.ForceRebuild3(false);
        }


        // --- NEW METHODS BELOW ---

        // Create a reference offset plane in an existing component.
        // Args: componentPath=full path to file; planeName=standard plane name (Front, Top, Right); offsetDistance=offset in meters.
        static void CreateReferenceOffsetPlane(SldWorks swApp, string componentPath, string planeName, double offsetDistance)
        // Begin CreateReferenceOffsetPlane method scope.
        {
            // Disable "Input Dimension Value" to avoid pauses during interactions.
            bool oldInputDimVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // Open the document if not already active; using OpenDoc6 for robustness.
            // Args: componentPath=file path; swDocumentTypes_e.swDocPART=assume part; swOpenDocOptions_Silent=no UI; ""=config; ref errors/warnings.
            int errors = 0, warnings = 0;
            ModelDoc2 swModel = swApp.OpenDoc6(componentPath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);

            // Verify document opened successfully.
            if (swModel == null)
            {
                // Restore preference if we exit early.
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);
                return;
            }

            // Normalize plane name to standard SOLIDWORKS names if user passes just "Front".
            string standardPlaneName = planeName;
            if (!planeName.Contains("Plane")) standardPlaneName += " Plane";

            // Select the standard plane to offset from.
            // Args: standardPlaneName=entity name; "PLANE"=entity type; 0,0,0=point; false=no append; 0=mark; null=callout; 0=options.
            bool status = swModel.Extension.SelectByID2(standardPlaneName, "PLANE", 0, 0, 0, false, 0, null, 0);

            // Create the offset plane if selection succeeded.
            if (status)
            {
                // CreatePlaneAtOffset3: Creates a reference plane.
                // Args: offsetDistance=distance in meters; false=normal direction; false=default options.
                Feature newPlane = (Feature)swModel.CreatePlaneAtOffset3(offsetDistance, false, false);

                // Rename the new plane for clarity.
                if (newPlane != null)
                {
                    // Naming convention: OriginalName + "_Offset" (e.g., "Front_Offset").
                    newPlane.Name = planeName + "_Offset";
                }
            }

            // Restore the original preference setting.
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            // Rebuild to update feature tree.
            swModel.ForceRebuild3(false);

            // Save the changes to the file.
            // Args: swSaveAsOptions_Silent=no prompts.
            swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
        }

        // Create a rectangular part centered at the origin with fully defined sketch and mid-plane extrusion.
        // Args: partName=filename; x=width (mm); y=depth (mm); z=height (mm); outputFolder=save location.
        static void CreateCenteredRectangularPart(SldWorks swApp, string partName, double x, double y, double z, string outputFolder)
        // Begin CreateCenteredRectangularPart method scope.
        {
            // Get default part template.
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            // Create new part document.
            ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);

            // Select Top Plane for the sketch.
            swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);

            // Create Center Rectangle: Ensures symmetry about origin (0,0).
            // Args: 0,0,0=Center Point; x/2000,y/2000,0=Corner Point (half dimensions in meters).
            // Note: CreateCenterRectangle automatically adds centerlines and symmetry relations.
            swModel.SketchManager.CreateCenterRectangle(0, 0, 0, x / 2000, y / 2000, 0);

            // Disable "Input Dimension Value" to prevent popup dialogs.
            bool oldInputDimVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // Add Smart Dimensions to fully define (constrain) the sketch.
            // 1. Dimension Width (X).
            swModel.ClearSelection2(true); // Clear created rectangle entities.
            // Select the top/bottom horizontal line (which is along X-axis, positioned at Z = +/- depth).
            // Top Plane layout: X is Model X, Y is Model Z. Sketch plane is at Model Y=0.
            // We select a point on the edge: X=0, Y=0, Z=(y/2000).
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, 0, (y / 2000), false, 0, null, 0);
            // Add dimension text.
            swModel.AddDimension2(0, 0, (y / 2000) + 0.02);

            // 2. Dimension Depth (Z).
            swModel.ClearSelection2(true);
            // Select the right/left vertical line (which is along Z-axis, positioned at X = +/- width).
            // We select a point on the edge: X=(x/2000), Y=0, Z=0.
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", (x / 2000), 0, 0, false, 0, null, 0);
            // Add dimension text.
            swModel.AddDimension2((x / 2000) + 0.02, 0, 0);

            // Restore the original preference setting.
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            // Setup extrusion feature.
            // Use Mid-Plane extrusion so planes remain centered in the part.
            // FeatureExtrusion2 Args:
            // ...
            // 4) swEndCondMidPlane (6) = End Condition.
            // 6) z/1000 = Depth (total thickness).
            swModel.FeatureManager.FeatureExtrusion2(true, false, false, (int)swEndConditions_e.swEndCondMidPlane, 0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

            // Save the part.
            string fullPath = Path.Combine(outputFolder, partName + ".SLDPRT");
            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);

            // Close the document to free memory.
            // swApp.CloseDoc(partName);
        }

        // Create a circular (cylindrical) part centered at origin with fully defined sketch and mid-plane extrusion.
        // Args: partName=filename; diameter=mm; z=height mm; outputFolder=save location.
        static void CreateCenteredCircularPart(SldWorks swApp, string partName, double diameter, double z, string outputFolder)
        // Begin CreateCenteredCircularPart method scope.
        {
            // Get default template and create new part.
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);

            // Select Top Plane.
            swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);

            // Create Circle at Origin.
            // Args: 0,0,0=Center; radius=diameter/2000 (mm to m radius).
            swModel.SketchManager.CreateCircleByRadius(0, 0, 0, diameter / 2000);

            // Disable "Input Dimension Value" to prevent popup dialogs.
            bool oldInputDimVal = swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);

            // Add Smart Dimension for Diameter.
            // Select the circle arc.
            // Args: diameter/2000, 0, 0 = approximately a point on the circumference.
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", diameter / 2000, 0, 0, false, 0, null, 0);
            // Add dimension text.
            swModel.AddDimension2((diameter / 2000) + 0.02, 0.02, 0);

            // Restore the original preference setting.
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, oldInputDimVal);

            // Extrude Mid-Plane.
            // Args: swEndCondMidPlane (6) ensures the cylinder is centered on the sketch plane.
            swModel.FeatureManager.FeatureExtrusion2(true, false, false, (int)swEndConditions_e.swEndCondMidPlane, 0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

            // Save the part.
            string fullPath = Path.Combine(outputFolder, partName + ".SLDPRT");
            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);

            // Close the document.
            // swApp.CloseDoc(partName);
        }
    }
}


