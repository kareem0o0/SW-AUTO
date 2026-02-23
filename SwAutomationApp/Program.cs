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
            string partAPath = partBuilder.GeneratePart(partAParameters, outFolder);
            string partBPath = partBuilder.GeneratePart(partBParameters, outFolder);

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
            // sketchBuilder.CreateRectangle(SketchRectangleType.Center, 0, 0, 40, 20);
            // sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 8);
            // sketchBuilder.CreateLine(SketchLineType.Centerline, -40, 0, 40, 0);
            // sketchBuilder.CreateLine(SketchLineType.Construction, -25, -20, -25, 20);
            // sketchBuilder.CreateLine(SketchLineType.Standard, -40, -20, 40, -20);
            // sketchBuilder.CreateLine(SketchLineType.Standard, -40, 20, 40, 20);
            // sketchBuilder.CreatePoint(SketchPointType.Standard, 0, 0);
            // sketchBuilder.CreatePoint(
            //     SketchPointType.Midpoint,
            //     0,
            //     -20,
            //     new SketchEntityReference("SKETCHSEGMENT", 0, -20));
            // sketchBuilder.ApplySketchRelation(
            //     SketchRelationType.Parallel,
            //     new SketchEntityReference("SKETCHSEGMENT", 0, -20),
            //     new SketchEntityReference("SKETCHSEGMENT", 0, 20));
            // sketchBuilder.CreateSketchFillet(
            //     5,
            //     new SketchEntityReference("SKETCHSEGMENT", -40, -20),
            //     new SketchEntityReference("SKETCHSEGMENT", -40, 20));
            // sketchBuilder.CreateSketchChamfer(
            //     SketchChamferMode.DistanceAngle,
            //     4,
            //     45,
            //     new SketchEntityReference("SKETCHSEGMENT", 40, -20),
            //     new SketchEntityReference("SKETCHSEGMENT", 40, 20));
            // sketchBuilder.AddDimension(0, 25, 0);
            // sketchBuilder.EndSketch();
            // sketchBuilder.Extrude(12, midPlane: false);
            // sketchBuilder.SaveAndClose(closeDocument: false);

        }
        catch (Exception ex)
        {
            // Main handles critical failures that bubble from lower-level classes.
            Console.WriteLine("Fatal error: " + ex.Message);
        }
    }
}
