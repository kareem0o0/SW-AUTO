using System; // Import base .NET types like Action, Exception, and Console.
using System.IO; // Import file system APIs for directory checks and path combining.

namespace SwAutomation; // Keep this class in the same project namespace.

public static class Examples // Expose reusable runnable examples without needing an instance.
{
    // Store a single default output folder so all examples write to the same location.
    private const string DefaultOutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";

    // Run the current default scenario: create two sample parts then mate them with custom face pairs.
    public static void RunCustomFacePairsCurrent()
    {
        // Open a SOLIDWORKS session and execute this example body with prepared builders.
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = sketchBuilder; // Mark sketchBuilder as intentionally unused in this example.
            GenerateSampleParts(partBuilder, outFolder, out string partAPath, out string partBPath); // Build Part_A and Part_B and return their saved file paths.

            // Only proceed if both part paths are valid.
            if (!string.IsNullOrWhiteSpace(partAPath) && !string.IsNullOrWhiteSpace(partBPath))
            {
                // Define custom mate mapping between faces of part A and part B.
                FaceMatePair[] customPairs =
                {
                    new FaceMatePair(FaceName.Top, FaceName.Bottom), // Mate top face of A to bottom face of B.
                    new FaceMatePair(FaceName.Left, FaceName.Right) // Mate left face of A to right face of B.
                };

                // Generate assembly using the explicit custom face recipe and save with this file name.
                assemblyBuilder.GenerateAssemblyByCustomFacePairs(
                    partAPath,
                    partBPath,
                    outFolder,
                    customPairs,
                    "Final_Assembly_CustomFaceMates.SLDASM");
            }
            else
            {
                // Inform user that assembly step did not run because part generation failed.
                Console.WriteLine("Assembly skipped: one or more parts were not generated successfully.");
            }
        });
    }

    // Run plane-based assembly mating after generating the same sample parts.
    public static void RunAssemblyByPlanes()
    {
        // Execute this example inside a managed SOLIDWORKS session.
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = sketchBuilder; // Mark sketchBuilder as intentionally unused in this example.
            GenerateSampleParts(partBuilder, outFolder, out string partAPath, out string partBPath); // Create source parts first.

            // Guard against missing outputs before calling assembly APIs.
            if (!string.IsNullOrWhiteSpace(partAPath) && !string.IsNullOrWhiteSpace(partBPath))
            {
                assemblyBuilder.GenerateAssembly(partAPath, partBPath, outFolder); // Build assembly using plane mates.
            }
            else
            {
                Console.WriteLine("Assembly skipped: one or more parts were not generated successfully."); // Report skip reason.
            }
        });
    }

    // Run default face-mate recipe after generating sample parts.
    public static void RunAssemblyByDefaultFaceMates()
    {
        // Execute this example inside a managed SOLIDWORKS session.
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = sketchBuilder; // Mark sketchBuilder as intentionally unused in this example.
            GenerateSampleParts(partBuilder, outFolder, out string partAPath, out string partBPath); // Create source parts first.

            // Continue only when both generated part paths are usable.
            if (!string.IsNullOrWhiteSpace(partAPath) && !string.IsNullOrWhiteSpace(partBPath))
            {
                assemblyBuilder.GenerateAssemblyByFaces(partAPath, partBPath, outFolder); // Use built-in default face mapping.
            }
            else
            {
                Console.WriteLine("Assembly skipped: one or more parts were not generated successfully."); // Report skip reason.
            }
        });
    }

    // Run a variant custom face-mate mapping to test another orientation recipe.
    public static void RunAssemblyByCustomFaceMappingVariant()
    {
        // Execute this example inside a managed SOLIDWORKS session.
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = sketchBuilder; // Mark sketchBuilder as intentionally unused in this example.
            GenerateSampleParts(partBuilder, outFolder, out string partAPath, out string partBPath); // Create source parts first.

            // Define alternate custom face pairs.
            FaceMatePair[] customPairsVariant =
            {
                new FaceMatePair(FaceName.Back, FaceName.Front), // Mate back of A to front of B.
                new FaceMatePair(FaceName.Left, FaceName.Right), // Mate left of A to right of B.
                new FaceMatePair(FaceName.Top, FaceName.Bottom) // Mate top of A to bottom of B.
            };

            // Run custom pair assembly only when source parts are available.
            if (!string.IsNullOrWhiteSpace(partAPath) && !string.IsNullOrWhiteSpace(partBPath))
            {
                assemblyBuilder.GenerateAssemblyByCustomFacePairs(
                    partAPath,
                    partBPath,
                    outFolder,
                    customPairsVariant,
                    "Final_Assembly_CustomVariant.SLDASM"); // Save variant assembly with dedicated name.
            }
            else
            {
                Console.WriteLine("Assembly skipped: one or more parts were not generated successfully."); // Report skip reason.
            }
        });
    }

    // Create only centered stock parts and skip all assembly operations.
    public static void RunCenteredStockPartsOnly()
    {
        // Execute this example inside a managed SOLIDWORKS session.
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = assemblyBuilder; // Mark assemblyBuilder as intentionally unused in this example.
            _ = sketchBuilder; // Mark sketchBuilder as intentionally unused in this example.
            partBuilder.CreateCenteredRectangularPart("RectPart_Centered", 120, 60, 20, outFolder); // Create centered rectangular part.
            partBuilder.CreateCenteredCircularPart("CircPart_Centered", 50, 100, outFolder); // Create centered circular part.
        });
    }

    // Create a fresh part and add two offset planes from the right plane.
    public static void RunCreatePartWithOffsetPlanes()
    {
        // Execute this example inside a managed SOLIDWORKS session.
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = assemblyBuilder; // Mark assemblyBuilder as intentionally unused in this example.
            _ = sketchBuilder; // Mark sketchBuilder as intentionally unused in this example.
            _ = outFolder; // Mark output folder as intentionally unused in this example.
            partBuilder.CreatePartWithOffsetPlanes(0.05); // Create two offset planes at 0.05 m (50 mm).

            // Start-part + extrude-existing-sketch flow example:
            // ModelDoc2 partModel = partBuilder.GeneratePart("StartedPart_1", SketchPlaneName.Front);
            // partModel.SketchManager.CreateCenterRectangle(0, 0, 0, 0.02, 0.01, 0); // meters in raw COM call (20x10 mm)
            // string startedPartPath = partBuilder.ExtrudeExistingSketchAndSave(partModel, "StartedPart_1", outFolder, 12, midPlane: false);
        });
    }

    // Open an existing part and add multiple reference offset planes.
    public static void RunAddOffsetPlanesToExistingPart()
    {
        // Execute this example inside a managed SOLIDWORKS session.
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = assemblyBuilder; // Mark assemblyBuilder as intentionally unused in this example.
            _ = sketchBuilder; // Mark sketchBuilder as intentionally unused in this example.
            string rectPartPath = Path.Combine(outFolder, "RectPart_Centered.SLDPRT"); // Build full path to existing part.
            partBuilder.CreateReferenceOffsetPlane(rectPartPath, "Front", 0.10); // Add 100 mm offset from Front plane.
            partBuilder.CreateReferenceOffsetPlane(rectPartPath, "Top", 0.05); // Add 50 mm offset from Top plane.
            partBuilder.CreateReferenceOffsetPlane(rectPartPath, "Right", 0.02); // Add 20 mm offset from Right plane.

            // Part mirror example:
            // partBuilder.MirrorPartFeature(
            //     rectPartPath,
            //     new SketchEntityReference("PLANE", 0, 0, 0, "Front Plane"),
            //     new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"));

            // Part linear pattern example:
            // partBuilder.CreatePartLinearPattern(
            //     rectPartPath,
            //     new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"),
            //     new SketchEntityReference("EDGE", 0, 0, 0),
            //     4,
            //     20);

            // Part circular pattern example:
            // partBuilder.CreatePartCircularPattern(
            //     rectPartPath,
            //     new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"),
            //     new SketchEntityReference("AXIS", 0, 0, 0),
            //     6,
            //     60);
        });
    }

    public static void RunSketchWorkflow()
    {
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = partBuilder;
            _ = assemblyBuilder;

            sketchBuilder.BeginPartSketch("ComprehensiveSketchPart", outFolder, SketchPlaneName.Front);
            
            sketchBuilder.CreateRectangle(SketchRectangleType.Center, 0, 0, 50, 50);

            
            sketchBuilder.CreateLine(SketchLineType.Centerline, -60, 0, 60, 0);
            sketchBuilder.CreateLine(SketchLineType.Construction, 0, -60, 0, 60);
            
            sketchBuilder.CreateLine(SketchLineType.Standard, 60, -20, 80, -20);
            sketchBuilder.CreateLine(SketchLineType.Standard, 80, -20, 80, 20);
            sketchBuilder.CreateLine(SketchLineType.Standard, 80, 20, 60, 20);
            sketchBuilder.CreateLine(SketchLineType.Standard, 60, 20, 60, -20);
            
            sketchBuilder.CreatePoint(-25, 25);

            sketchBuilder.ApplySketchRelation(SketchRelationType.Vertical, new SketchEntityReference(80, 0));

            // Sketch mirror example:
            // sketchBuilder.MirrorSketchEntities(
            //     new SketchEntityReference("SKETCHSEGMENT", 0, 0),
            //     new SketchEntityReference("SKETCHSEGMENT", 60, -20),
            //     new SketchEntityReference("SKETCHSEGMENT", 80, -20),
            //     new SketchEntityReference("SKETCHSEGMENT", 80, 20),
            //     new SketchEntityReference("SKETCHSEGMENT", 60, 20));

            // Sketch linear pattern example:
            // sketchBuilder.CreateLinearSketchPattern(
            //     3, 1, 20, 0, 0, 90, "",
            //     new SketchEntityReference("SKETCHSEGMENT", 60, -20));

            // Sketch circular pattern example:
            // sketchBuilder.CreateCircularSketchPattern(
            //     40, 180, 6, 30, false, "",
            //     new SketchEntityReference("SKETCHSEGMENT", 60, -20));

            sketchBuilder.CreateSketchFillet(5, new SketchEntityReference(-50, 50));
            sketchBuilder.CreateSketchChamfer(SketchChamferMode.DistanceDistance, 5, new SketchEntityReference(50, 50));
            
            sketchBuilder.AddDimension(0, 50, 0, 60, 0);
            sketchBuilder.AddDimension(60, 0, 80, 0, 70, 30, 0);
            
            sketchBuilder.EndSketch();
            sketchBuilder.Extrude(15, midPlane: true); // Circular results auto-create Face_* helper planes.
            sketchBuilder.SaveAndClose(closeDocument: false);
        });
    }

    // Shared runner that opens session, builds helper services, ensures output folder, and executes an example action.
    private static void RunWithSession(Action<PartBuilder, AssemblyBuilder, SketchBuilder, string> example)
    {
        string outFolder = DefaultOutputFolder; // Use configured default output folder.

        try // Wrap session lifecycle and example execution with global error handling.
        {
            Console.WriteLine("Connecting to SOLIDWORKS..."); // Inform user before COM session creation.

            using SwSession session = new SwSession(visible: true); // Start visible SOLIDWORKS session and auto-dispose on scope exit.

            PartBuilder partBuilder = new PartBuilder(session.Application); // Create part automation helper.
            AssemblyBuilder assemblyBuilder = new AssemblyBuilder(session.Application); // Create assembly automation helper.
            SketchBuilder sketchBuilder = new SketchBuilder(session.Application); // Create sketch automation helper.

            if (!Directory.Exists(outFolder)) // Check whether output folder already exists.
            {
                Directory.CreateDirectory(outFolder); // Create output folder when missing.
            }

            example(partBuilder, assemblyBuilder, sketchBuilder, outFolder); // Execute the selected example body.
        }
        catch (Exception ex) // Catch unhandled failures from setup or example execution.
        {
            Console.WriteLine("Fatal error: " + ex.Message); // Print top-level fatal error message.
        }
    }

    // Helper that generates the standard two sample parts used by multiple assembly examples.
    private static void GenerateSampleParts(PartBuilder partBuilder, string outFolder, out string partAPath, out string partBPath)
    {
        PartParameters partAParameters = new PartParameters("Part_A", 100, 50, 10, 5, 3, SketchPlaneName.Front); // Define Part_A dimensions and hole pattern.
        PartParameters partBParameters = new PartParameters("Part_B", 80, 80, 20, 4, 2, SketchPlaneName.Top); // Define Part_B dimensions and hole pattern.

        partAPath = partBuilder.GenerateRectangularPartWithHoles(partAParameters, outFolder); // Create Part_A and capture saved path.
        partBPath = partBuilder.GenerateRectangularPartWithHoles(partBParameters, outFolder); // Create Part_B and capture saved path.
    }
    public static void Playground()
    {
        // Execute this example inside a managed SOLIDWORKS session.
        RunWithSession((partBuilder, assemblyBuilder, sketchBuilder, outFolder) =>
        {
            _ = partBuilder; // Mark partBuilder as intentionally unused in this sketch-only workflow.
            _ = assemblyBuilder; // Mark assemblyBuilder as intentionally unused in this sketch-only workflow.

            sketchBuilder.BeginPartSketch("SketchPart_1", outFolder, SketchPlaneName.Front); // Create part and start sketch on Front plane.
            sketchBuilder.CreateRectangle(SketchRectangleType.Center, 0, 0, 40, 20); // Add center rectangle (mm coordinates).
            
            // Add intersecting lines and demonstrate trim without specifying z (uses probe fallback).
            sketchBuilder.CreateLine(SketchLineType.Standard, 20, -30, 20, 30); // Vertical line crossing sketch plane.
            sketchBuilder.CreateLine(SketchLineType.Standard, -30, 0, 40, 0); // Horizontal line across origin.
            // Trim the vertical line at the horizontal intersection.
            // Provide explicit Z (0 mm) and use closest-trim mode (0).
            sketchBuilder.CreateSketchTrim(20, 0, 0, 0);
            //sketchBuilder.CreateLine(SketchLineType.Centerline, -40, 0, 40, 0); // Add horizontal centerline.
            //sketchBuilder.CreateLine(SketchLineType.Construction, -25, -20, -25, 20); // Add vertical construction line.
            //sketchBuilder.CreateLine(SketchLineType.Standard, -40, -20, 40, -20); // Add lower standard line.
            //sketchBuilder.CreateLine(SketchLineType.Standard, -40, 20, 40, 20); // Add upper standard line.
            //sketchBuilder.CreatePoint(0, 0); // Add standard point at origin.
            //sketchBuilder.CreatePoint(0, -20); // Add second standard point.
            //sketchBuilder.ApplySketchRelation(
            //    SketchRelationType.Parallel,
            //    new SketchEntityReference("SKETCHSEGMENT", 0, 10),
            //    new SketchEntityReference("SKETCHSEGMENT", -25, 0)); // Make the two segments parallel.
           // sketchBuilder.CreateSketchFillet(
              //  5,
              //  new SketchEntityReference("SKETCHSEGMENT", -40, -20),
              //  new SketchEntityReference("SKETCHSEGMENT", -40, 20)); // Add 5 mm sketch fillet between two segments.
            // Single-edge/point fillet mode (optional): only first reference is needed.
            sketchBuilder.CreateSketchFillet(3, new SketchEntityReference(-40, -20)); // Default type = SKETCHPOINT.
            sketchBuilder.CreateSketchChamfer(
                SketchChamferMode.DistanceDistance,
                4,
                new SketchEntityReference(40, -20)); // Equal-distance chamfer using one corner point (default type = SKETCHPOINT).
            // Two-distance chamfer example (requires both values):
            // sketchBuilder.CreateSketchChamfer(SketchChamferMode.DistanceDistance, 4, 6, new SketchEntityReference("SKETCHPOINT", 40, -20));
            // Single-edge/point chamfer mode (optional): only first reference is needed.
            // sketchBuilder.CreateSketchChamfer(SketchChamferMode.DistanceAngle, 2, 45, new SketchEntityReference("SKETCHPOINT", 40, -20));
            //sketchBuilder.AddDimension(-40, 20, 40, 20, 0, 30, 0); // AddDimension(firstX, firstY, secondX, secondY, textX, textY, textZ): dimension between points (-40,20) and (40,20), with text at (0,30,0).
            // Single-entity Smart Dimension example (line/circle/arc, inferred from one pick):
            sketchBuilder.AddDimension(40, 0, 50, 0, 0);
            sketchBuilder.EndSketch(); // Exit sketch edit mode.
            sketchBuilder.Extrude(12, midPlane: false); // Blind extrude sketch by 12 mm.
            sketchBuilder.SaveAndClose(closeDocument: false); // Save part and keep it open.
        });
    }
}


 // ---------------------------------------------------------------------
            // EXAMPLE: Directly calling native SOLIDWORKS API functions
            // You can always bypass SketchBuilder and use raw API methods directly 
            // inside Main for specific or complex operations.
            // ---------------------------------------------------------------------

                        /* Trim mode reference and usage examples:
                         - swSketchTrimClosest (0): Trim to closest. Select one segment, then call:
                                 _activeModel.Extension.SelectByID2("", "SKETCHSEGMENT", x/1000.0, y/1000.0, z/1000.0, false, 0, null, 0);
                                 sketchBuilder.CreateSketchTrim(xMm, yMm, zMm, 0);

                         - swSketchTrimCorner (1): Corner trim for two segments. Select two segments
                             (first append=false, second append=true), then call:
                                 sketchBuilder.CreateSketchTrim(xMm, yMm, zMm, 1);

                         - swSketchTrimTwoEntities (2): Trim first selected segment to second. Select
                             two segments using Mark=0 for pick points, then call:
                                 sketchBuilder.CreateSketchTrim(x1Mm, y1Mm, zMm, 2);

                         - swSketchTrimEntityPoint (3): Trim a single segment to a specific point on it.
                             Select the segment and call with the exact point on the segment:
                                 sketchBuilder.CreateSketchTrim(pxMm, pyMm, pzMm, 3);

                         - swSketchTrimEntities (4): Power trim multiple selected segments. Select segments
                             with Mark=2 and pick points, then call:
                                 // select segments with mark=2 using SelectByID2(..., mark:2)
                                 sketchBuilder.CreateSketchTrim(anyXmm, anyYmm, anyZmm, 4);

                         - swSketchTrimOutside (5): Trim away outside boundary. Select boundary and
                             target segments, then call:
                                 sketchBuilder.CreateSketchTrim(anyXmm, anyYmm, anyZmm, 5);

                         - swSketchTrimInside (6): Trim away inside boundary. Select boundary and
                             target segments, then call:
                                 sketchBuilder.CreateSketchTrim(anyXmm, anyYmm, anyZmm, 6);
                        */

            // Power-trim example removed per request.
            
            //sketchBuilder.DeleteSketchEntityAt(50, 100);
            //sketchBuilder.CreateLine(SketchLineType.Standard, 8.25, 300, 0.7, 300);
            //sketchBuilder.CreateLine(SketchLineType.Standard, 0.7, 300, 1.5, 310);
            //sketchBuilder.AddDimension( -8, 300, -8,280);
            // sketchBuilder.CreateLine(SketchLineType.Centerline, -40, 0, 40, 0);
            // sketchBuilder.CreateLine(SketchLineType.Construction, -25, -20, -25, 20);
            // sketchBuilder.CreateLine(SketchLineType.Standard, -40, -20, 40, -20);
            
            
            
            
/*
            // Start SOLIDWORKS session and keep UI visible for debugging.
            using SwSession session = new SwSession(visible: true);

            // Create specialized services that contain all automation logic.
            PartBuilder partBuilder = new PartBuilder(session.Application);
            AssemblyBuilder assemblyBuilder = new AssemblyBuilder(session.Application);
            SketchBuilder sketchBuilder = new SketchBuilder(session.Application);
            _ = sketchBuilder; // Keep module ready for optional sketch-workflow examples below.

            // Ensure output folder exists before saving files.
            if (!Directory.Exists(outFolder)) Directory.CreateDirectory(outFolder);

            // Build parameter objects instead of long string argument lists.
            // Sketch plane options for GeneratePart:
            // - SketchPlaneName.Front (default)
            // - SketchPlaneName.Top
            // - SketchPlaneName.Right
            // - SketchPlaneName.Side (alias of Right)
            PartParameters partAParameters = new PartParameters("Part_A", 100, 50, 10, 5, 3, SketchPlaneName.Front);
            PartParameters partBParameters = new PartParameters("Part_B", 80, 80, 20, 4, 2, SketchPlaneName.Top);

            // Generate both parts.
            string partAPath = partBuilder.GenerateRectangularPartWithHoles(partAParameters, outFolder);
            string partBPath = partBuilder.GenerateRectangularPartWithHoles(partBParameters, outFolder);

            // Continue to assembly only when both parts exist.
            if (!string.IsNullOrWhiteSpace(partAPath) && !string.IsNullOrWhiteSpace(partBPath))
                 
             - swSketchTrimClosest (0): Trim to closest. Call with mode-first and point:
                 sketchBuilder.CreateSketchTrim(0, xMm, yMm, zMm);
                FaceMatePair[] customPairs =
                {
                    new FaceMatePair(FaceName.Top, FaceName.Bottom),
                    new FaceMatePair(FaceName.Left, FaceName.Right),
                    // Mixed mode example (point selector on A, named face on B):
                    // new FaceMatePair(new FaceSelection(50, 0, 5), new FaceSelection(FaceName.Right)),

                };

                // Build assembly by explicit face-pair recipe.
                assemblyBuilder.GenerateAssemblyByCustomFacePairs(partAPath, partBPath, outFolder, customPairs, "Final_Assembly_CustomFaceMates.SLDASM");
            }
            else
            {
                Console.WriteLine("Assembly skipped: one or more parts were not generated successfully.");
            }
*/
            // ---------------------------------------------------------------------
            // Usage examples (uncomment any block you want to run)
            // ---------------------------------------------------------------------

            // EXAMPLE 1: Plane-based mating using reference planes (Front/Top).
            // if (!string.IsNullOrWhiteSpace(partAPath) && !string.IsNullOrWhiteSpace(partBPath))
            // {
            //     assemblyBuilder.GenerateAssembly(partAPath, partBPath, outFolder);
            // }

            // EXAMPLE 2: Default face-mate recipe.
            // Recipe = Top->Bottom, Right->Right, Front->Front.
            // if (!string.IsNullOrWhiteSpace(partAPath) && !string.IsNullOrWhiteSpace(partBPath))
            // {
            //     assemblyBuilder.GenerateAssemblyByFaces(partAPath, partBPath, outFolder);
            // }

            // EXAMPLE 3: Custom face mapping with different orientation.
            // FaceMatePair[] customPairsVariant = new[]
            // {
            //     new FaceMatePair(FaceName.Back, FaceName.Front),
            //     new FaceMatePair(FaceName.Left, FaceName.Right),
            //     new FaceMatePair(FaceName.Top, FaceName.Bottom)
            // };
            // if (!string.IsNullOrWhiteSpace(partAPath) && !string.IsNullOrWhiteSpace(partBPath))
            // {
            //     assemblyBuilder.GenerateAssemblyByCustomFacePairs(
            //         partAPath,
            //         partBPath,
            //         outFolder,
            //         customPairsVariant,
            //         "Final_Assembly_CustomVariant.SLDASM");
            // }

            // EXAMPLE 4: Create only centered stock parts (no assembly).
            // partBuilder.CreateCenteredRectangularPart("RectPart_Centered", 120, 60, 20, outFolder);
            // partBuilder.CreateCenteredCircularPart("CircPart_Centered", 50, 100, outFolder);

            // EXAMPLE 5: Create a new part with two offset planes from Right Plane.
            // Input in meters: 0.05 = 50 mm.
            // partBuilder.CreatePartWithOffsetPlanes(0.05);

            // EXAMPLE 6: Add offset planes into an existing part file.
            // string rectPartPath = Path.Combine(outFolder, "RectPart_Centered.SLDPRT");
            // partBuilder.CreateReferenceOffsetPlane(rectPartPath, "Front", 0.10); // 100 mm
            // partBuilder.CreateReferenceOffsetPlane(rectPartPath, "Top", 0.05);   // 50 mm
            // partBuilder.CreateReferenceOffsetPlane(rectPartPath, "Right", 0.02); // 20 mm

            // EXAMPLE 7: Step-by-step sketch module workflow (new).
            // sketchBuilder.BeginPartSketch("SketchPart_1", outFolder, SketchPlaneName.Front);
            // Existing-part sketch start examples:
            // ModelDoc2 existingPart = (ModelDoc2)session.Application.ActiveDoc;
            // sketchBuilder.BeginSketch(SketchPlaneName.Top, existingPart); // Start on main datum plane
            // sketchBuilder.BeginSketch(new SketchEntityReference("FACE", 40, 0, 10), existingPart); // Start on planar part surface
            // sketchBuilder.CreateRectangle(SketchRectangleType.Center, 0, 0, 40, 20);
            // sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 8);
            // sketchBuilder.CreateLine(SketchLineType.Centerline, -40, 0, 40, 0);
            // sketchBuilder.CreateLine(SketchLineType.Construction, -25, -20, -25, 20);
            // sketchBuilder.CreateLine(SketchLineType.Standard, -40, -20, 40, -20);
            // sketchBuilder.CreateLine(SketchLineType.Standard, -40, 20, 40, 20);
            // sketchBuilder.CreatePoint(0, 0);
            // sketchBuilder.CreatePoint(0, -20);
            // sketchBuilder.ApplySketchRelation(
            //     SketchRelationType.Parallel,
            //     new SketchEntityReference("SKETCHSEGMENT", 0, -20),
            //     new SketchEntityReference("SKETCHSEGMENT", 0, 20));
            // sketchBuilder.MirrorSketchEntities(
            //     new SketchEntityReference("SKETCHSEGMENT", 0, 0),
            //     new SketchEntityReference("SKETCHSEGMENT", -40, -20));
            // sketchBuilder.CreateLinearSketchPattern(
            //     3, 1, 20, 0, 0, 90, "",
            //     new SketchEntityReference("SKETCHSEGMENT", -40, -20));
            // sketchBuilder.CreateCircularSketchPattern(
            //     40, 180, 6, 30, false, "",
            //     new SketchEntityReference("SKETCHSEGMENT", -40, -20));
            // sketchBuilder.CreateSketchFillet(
            //     5,
            //     new SketchEntityReference("SKETCHSEGMENT", -40, -20),
            //     new SketchEntityReference("SKETCHSEGMENT", -40, 20));
            // Optional single-edge/point mode:
            // sketchBuilder.CreateSketchFillet(3, new SketchEntityReference("SKETCHPOINT", -40, -20));
            // sketchBuilder.CreateSketchChamfer(
            //     SketchChamferMode.DistanceAngle,
            //     4,
            //     45,
            //     new SketchEntityReference("SKETCHSEGMENT", 40, -20),
            //     new SketchEntityReference("SKETCHSEGMENT", 40, 20));
            // Optional single-edge/point mode:
            // sketchBuilder.CreateSketchChamfer(SketchChamferMode.DistanceAngle, 2, 45, new SketchEntityReference("SKETCHPOINT", 40, -20));
            // Optional equal-distance chamfer overload:
            // sketchBuilder.CreateSketchChamfer(SketchChamferMode.DistanceDistance, 4, new SketchEntityReference("SKETCHPOINT", 40, -20));
            // sketchBuilder.AddDimension(-40, 20, 40, 20, 0, 30, 0);
            // Optional single-entity Smart Dimension:
            // sketchBuilder.AddDimension(40, 20, 20, 35, 0);
            // sketchBuilder.EndSketch();
            // sketchBuilder.Extrude(12, midPlane: false);
            // sketchBuilder.SaveAndClose(closeDocument: false);

            // EXAMPLE 8: Part-level mirror/linear/circular patterns (after part is created).
            // partBuilder.MirrorPartFeature(
            //     partAPath,
            //     new SketchEntityReference("PLANE", 0, 0, 0, "Front Plane"),
            //     new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"));
            // partBuilder.CreatePartLinearPattern(
            //     partAPath,
            //     new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"),
            //     new SketchEntityReference("EDGE", 0, 0, 0),
            //     4,
            //     20);
            // partBuilder.CreatePartCircularPattern(
            //     partAPath,
            //     new SketchEntityReference("BODYFEATURE", 0, 0, 0, "Boss-Extrude1"),
            //     new SketchEntityReference("AXIS", 0, 0, 0),
            //     6,
            //     60);

            // ---------------------------------------------------------------------
            // Examples.cs methods (commented call shortcuts)
            // ---------------------------------------------------------------------
            // Examples.RunCustomFacePairsCurrent();
            // Examples.RunAssemblyByPlanes();
            // Examples.RunAssemblyByDefaultFaceMates();
            // Examples.RunAssemblyByCustomFaceMappingVariant();
            // Examples.RunCenteredStockPartsOnly();
            // Examples.RunCreatePartWithOffsetPlanes();
            // Examples.RunAddOffsetPlanesToExistingPart();
            // Examples.RunSketchWorkflow();
            // Examples.Playground();