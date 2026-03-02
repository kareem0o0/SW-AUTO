

/*// Import base .NET types; alternative: remove if unused to speed build slightly.
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
            

            sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 990/2.0); // args: (circleType, centerXmm, centerYmm, radiusMm)
            sketchBuilder.CreateCircle(SketchCircleType.CenterRadius, 0, 0, 640/2.0); // args: (circleType, centerXmm, centerYmm, radiusMm)
            sketchBuilder.AddDimension(990/2.0, 0, 990/2.0+20, 20 );//(pickXmm, pickYmm, textXmm, textYmm, valueMm?) 
            sketchBuilder.AddDimension(640/2.0, 0, 640/2.0+20, 20); // radius
            sketchBuilder.EndSketch();
            sketchBuilder.Extrude(8, midPlane: false, isCut: false); // args: (depthMm, midPlane, isCut)
            sketchBuilder.BeginSketch(new SketchEntityReference("FACE", 350, 0, 8));
            sketchBuilder.DisableSketchInference();
            sketchBuilder.CreateRectangle(SketchRectangleType.Center, 0, 320-20+115.2/2, 15/7.0, 320+85.2);
            sketchBuilder.AddDimension(-15/7.0, 320-20, -15/7.0+10, 320-20);
            sketchBuilder.AddDimension(15/7.0, 320+85.2, 15/7.0+10, 320+85.2);
            
            sketchBuilder.CreatePoint(0, 0);
            sketchBuilder.CreatePoint(0, 320-20+115.2/2);
            //sketchBuilder.EnableSketchInference();
            sketchBuilder.ApplySketchRelation(SketchRelationType.Vertical,
              new SketchEntityReference("SKETCHPOINT", 0, 357.6),  // rectangle center
             new SketchEntityReference("SKETCHPOINT", 0, 0));     // origin point
            //sketchBuilder.ApplySketchRelation(SketchRelationType.Coincident, 0, 100, "SKETCHSEGMENT", -8.25, 350, "SKETCHSEGMENT");
            //sketchBuilder.CreateLine(SketchLineType.Standard, 15/7.0+20, 320+0.7, 15/7.0+3, 320+5.7);
            



            //sketchBuilder.CreateSketchTrim(1, 300);
            //sketchBuilder.CreateSketchTrim(-8.25, 300);
            ///sketchBuilder.CreateSketchTrim(8.25, 300);
            ///sketchBuilder.EnableSketchInference();
            //sketchBuilder.AddDimension(-1500.25, 300, -1300, 300, -1400, 250, 200.25);//args: (firstXmm, firstYmm, secondXmm, secondYmm, textXmm, textYmm, valueMm?)
            //sketchBuilder.CreateLine(SketchLineType.Standard, 50, 100, 150, 100);
            //sketchBuilder.ApplySketchRelation(SketchRelationType.Parallel, 105, 100, "SKETCHSEGMENT", -8.25, 350, "SKETCHSEGMENT");
            //sketchBuilder.ApplySketchRelation(SketchRelationType.Vertical, new SketchEntityReference("SKETCHSEGMENT", 50, 50)); // args: (relationType, entityRef)
            //sketchBuilder.CreateSketchTrim(1, 300); // args: (Xmm,Ymm))
            
        }
        catch (Exception ex)
        {
            // Main handles critical failures that bubble from lower-level classes.
            Console.WriteLine("Fatal error: " + ex.Message);
        }
    }
}
*/