// Import base .NET types; alternative: remove if unused to speed build slightly.
using System;
// Import file system helpers for folders and paths.
using System.IO;

// Removes platform warnings by specifying this code is Windows-only.
[assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")]

namespace SwAutomation;

// Program entry type; coordinates session and service classes.
public static class Program
{
    // App entry point; args can hold future command-line options.
    public static void Main(string[] args)
    {
        // Output folder where parts and assemblies are saved.
        string outFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";

        try
        {
            // Inform the user we're about to connect to SOLIDWORKS.
            Console.WriteLine("Connecting to SOLIDWORKS...");
            using SwSession session = new SwSession(visible: true);

            // Create specialized services that contain all automation logic.
            PartBuilder partBuilder = new PartBuilder(session.Application);
            AssemblyBuilder assemblyBuilder = new AssemblyBuilder(session.Application);
            SketchBuilder sketchBuilder = new SketchBuilder(session.Application);

            if (!Directory.Exists(outFolder)) Directory.CreateDirectory(outFolder);

            sketchBuilder.BeginPartSketch("StatorBleche", outFolder, SketchPlaneName.Front);
            sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 600); // x , y ,R
            sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 625/2.0); // x , y ,R
            sketchBuilder.AddDimension(600, 0, 600+20, 0, 0);
            sketchBuilder.AddDimension(625/2, 0, 625/2+20, 0, 0);
            sketchBuilder.EndSketch();
            sketchBuilder.Extrude(8, midPlane: false, isCut: false); 
            sketchBuilder.BeginSketch(new SketchEntityReference("FACE", 550, 0, 8));
            sketchBuilder.CreateLine(SketchLineType.Standard, 8.25, 300, 0.7, 300);
            sketchBuilder.CreateLine(SketchLineType.Standard, 0.7, 300, 1.5, 310);
            //sketchBuilder.CreateRectangle(SketchRectangleType.Corner, -8.25, 300, 8.25, 117 + 625/2.0);
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
            {
                // Example custom face mapping.
                // FaceName options:
                // Top, Bottom, Left, Right, Front, Back, Up, Down
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
        }
        catch (Exception ex)
        {
            // Main handles critical failures that bubble from lower-level classes.
            Console.WriteLine("Fatal error: " + ex.Message);
        }
    }
}
