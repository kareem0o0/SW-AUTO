// Import base .NET types; alternative: remove if unused to speed build slightly.
using System;
// Import file system helpers for folders and paths.
using System.IO;
// Import SOLIDWORKS interop types.
using SolidWorks.Interop.sldworks;
// Import SOLIDWORKS constant enums.
using SolidWorks.Interop.swconst;
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

            // ---------------------------------------------------------------------
            // Create a new part and begin sketching
            // ---------------------------------------------------------------------
            sketchBuilder.BeginPartSketch("StatorBleche", outFolder, SketchPlaneName0.Front);

            // IMPORTANT: Get the active document *after* BeginPartSketch creates the new part.
            // If you get this before BeginPartSketch, it might be null or point to a previous file.
            ModelDoc2 swModel = (ModelDoc2)session.Application.ActiveDoc;
            SelectionMgr selMgr = swModel.SelectionManager;

            sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 600); // x, y, R
            sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 625/2.0); // x, y, R
            sketchBuilder.CreateRectangle(SketchRectangleType.Corner, -1500.25, 300, -1300, 429.5);
            sketchBuilder.AddDimension(-1500.25, 300, -1300, 300, -1400, 250, 200.25);//args: (firstXmm, firstYmm, secondXmm, secondYmm, textXmm, textYmm, valueMm?)
            sketchBuilder.AddDimension(-1300, 300, -1300, 429.5, -1250, 364.75, 129.5);
            sketchBuilder.AddDimension(600, 0, 620, 20, 600);//(pickXmm, pickYmm, textXmm, textYmm, valueMm?) 
            sketchBuilder.AddDimension(312.5, 0, 340, 20, 312.5); // radius
            sketchBuilder.EndSketch();
            sketchBuilder.Extrude(8, midPlane: false, isCut: false); // args: (depthMm, midPlane, isCut)
            sketchBuilder.BeginSketch(new SketchEntityReference("FACE", 250, 0, 8));
            sketchBuilder.CreateRectangle(SketchRectangleType.Center, 0, 364.75, 8.25, 64.75);
            sketchBuilder.CreateSketchTrim(-8, 300, 8, 0); // args: (xMm, yMm, zMm, trimMode)
            sketchBuilder.CreateLine(SketchLineType.Standard, 50, 100, 100, 150); // args: (lineType, startXmm, startYmm, endXmm, endYmm)
            sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 800, 429.5, 625/2.0); // args: (circleType, centerXmm, centerYmm, radiusMm)
            sketchBuilder.ApplySketchRelation(SketchRelationType.Parallel, -8.25, 350, 8.25, 350); // args: (relationType, firstXmm, firstYmm, secondXmm, secondYmm)
            sketchBuilder.ApplySketchRelation(SketchRelationType.Vertical, new SketchEntityReference("SKETCHPOINT", -8.25, 350)); // args: (relationType, entityRef)
        }
        catch (Exception ex)
        {
            // Main handles critical failures that bubble from lower-level classes.
            Console.WriteLine("Fatal error: " + ex.Message);
        }
    }
}