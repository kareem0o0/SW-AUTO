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
    private const double MmToM = 0.001;
    private static double Mm(double mm) => mm * MmToM;

    // App entry point; args can hold future command-line options.
    [STAThread]
    public static void Main(string[] args)
    {
        // Output folder where parts and assemblies are saved.
        string outFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";

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
        double bottomYoffset = Mm(20.0);
        double angleInRadians = angleInDegrees * Math.PI / 180.0;
        double radiusMm = Mm(filletradius);

        SldWorks swApp = null;
        ModelDoc2 swModel = null;
        SketchManager swSketchManager = null;
        
        try
        {
            // Inform the user we're about to connect to SOLIDWORKS.
            Console.WriteLine("Connecting to SOLIDWORKS...");
            
            // Connect to SOLIDWORKS
            swApp = new SldWorks();
            swApp.Visible = true;
            Dimension swDim = null;
            DisplayDimension displayDim = null;
            
            if (!Directory.Exists(outFolder)) 
                Directory.CreateDirectory(outFolder);

            // ---------------------------------------------------------------------
            // Create a new part
            // ---------------------------------------------------------------------
            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            swModel = swApp.NewDocument(template, 0, 0, 0);
            
            if (swModel == null)
                throw new Exception("Failed to create new part");
            
            // Get SketchManager
            swSketchManager = swModel.SketchManager;
            SelectionMgr selMgr = swModel.SelectionManager;
            
            // Disable sketch inference (snapping)
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            
            // ---------------------------------------------------------------------
            // First Sketch: Front Plane
            // ---------------------------------------------------------------------
            // Select Front Plane
            bool selected = swModel.Extension.SelectByID2("Ebene vorne", "PLANE", 0, 0, 0, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select Front Plane");
            
            // Insert sketch
            swSketchManager.InsertSketch(true);
            
            // Create circles (values in meters)
            swSketchManager.CreateCircleByRadius(0, 0, 0, outerRadius);
            swSketchManager.CreateCircleByRadius(0, 0, 0, innerRadius);
            
            // Add dimensions for circles
            // First circle - diameter dimension (990mm)
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", outerRadius, 0, 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(outerRadius + Mm(20), Mm(20), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = Mm(outerDiameterMm);
            
            // Second circle - diameter dimension (640mm)
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", innerRadius, 0, 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(innerRadius + Mm(20), Mm(20), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = Mm(innerDiameterMm);
            
            // Extrude
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("Skizze1", "SKETCH", 0.4, 0, 0, false, 0, null, 0);
            swModel.FeatureManager.FeatureExtrusion2(
                true, false, false, 
                (int)swEndConditions_e.swEndCondBlind, 0, 
                plateThickness, 0,
                false, false, false, false, 
                0, 0, false, false, false, false, 
                true, true, true, 0, 0, false);
            
            // ---------------------------------------------------------------------
            // Second Sketch: On the top face of extrusion
            // ---------------------------------------------------------------------
            swModel.ClearSelection2(true);
            
            // Select the top face at Z=8mm
            selected = swModel.Extension.SelectByID2("", "FACE", (outerRadius + innerRadius) / 2.0, 0, plateThickness, false, 0, null, 0);
            if (!selected) throw new Exception("Could not select top face");
            
            // Insert new sketch on this face
            swSketchManager.InsertSketch(true);
            
            swSketchManager.CreateCenterRectangle(
                0, centerY, 0, 
                halfWidth, topY, 0);
            
            // Add width dimension
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(leftX + Mm(10), bottomY - Mm(10), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = halfWidth * 2;
            
            // Add height dimension
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, bottomY + Mm(50), 0, false, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(halfWidth + Mm(50), bottomY + Mm(50), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = topY - bottomY;
            
            
            // Add vertical relation between rectangle center and origin
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHPOINT",0, centerY, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0,  0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgVERTICALPOINTS2D");
            swModel.ClearSelection2(true);
            // Add dimensions from rectangle bottom to origin
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT",0, bottomY, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0,  0, 0, true, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY/2, 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = innerRadius;
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
            // Creating lines !
            // Capture the lines into variables
            SketchSegment line1 = (SketchSegment)swSketchManager.CreateLine(Mm(20), Mm(20), 0, Mm(50), Mm(50), 0);
            SketchSegment line2 = (SketchSegment)swSketchManager.CreateLine(Mm(50), Mm(50), 0, Mm(20), Mm(50), 0);

            // Later in your code, you can select line1 directly:
            
            swModel.Extension.SelectByID2("", "SKETCHPOINT", Mm(20),  Mm(20), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth,  bottomY + Mm(50), 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            
            swModel.Extension.SelectByID2("", "SKETCHPOINT", Mm(20),  Mm(50), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth,  bottomY + Mm(50), 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);
            
          
            //Add fillets
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth,  Mm(50), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth,  Mm(20), 0, true, 0, null, 0);
            SketchSegment slotfilletArc = swSketchManager.CreateFillet(radiusMm, (int)swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry);

            line2.Select4(false, null);
            swModel.SketchAddConstraints("sgHORIZONTAL");
            swModel.ClearSelection2(true);

            // Add dimensions from  slot profile
            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth,  Mm(50), 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth,  Mm(20), 0, true, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY-Mm(20), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = Mm(5.7);
            swModel.ClearSelection2(true);
            
    
            

            line2.Select4(false, null);
            line1.Select4(true, null);
            displayDim = (DisplayDimension)swModel.AddDimension2(0,Mm(30) , 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = angleInRadians;
            swModel.ClearSelection2(true);
            

            swModel.Extension.SelectByID2("", "SKETCHSEGMENT", 0,  bottomY, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth,  Mm(44.3), 0, true, 0, null, 0);
            displayDim = (DisplayDimension)swModel.AddDimension2(0, bottomY-Mm(20), 0);
            swDim = displayDim.GetDimension();
            swDim.SystemValue = Mm(-0.7);
            swModel.ClearSelection2(true);

            //Add fillets
            swModel.Extension.SelectByID2("", "SKETCHPOINT", halfWidth,  topY, 0, false, 0, null, 0);
            swSketchManager.CreateFillet(radiusMm,(int)swConstrainedCornerAction_e.swConstrainedCornerKeepGeometry);
            swModel.ClearSelection2(true);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", -halfWidth,  topY, 0, false, 0, null, 0);
            swSketchManager.CreateFillet(radiusMm,(int)swConstrainedCornerAction_e.swConstrainedCornerKeepGeometry);
            swModel.ClearSelection2(true);

        
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", halfWidth, topY-Mm(50), 0, false, 0, null, 0);
            swSketchManager.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimClosest, halfWidth, Mm(325.0), 0);

            
            // mirror the slot profile to the other side
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
            SketchSegment lineseg1 = (SketchSegment)swSketchManager.CreateCenterLine(0, 0, 0, 0, bottomY, 0);

            swModel.ClearSelection2(true);

            // For mirror operation, marks are critical [citation:2]
            // Mirror line/centerline needs mark 2
            SelectData mirrorData = selMgr.CreateSelectData();
            mirrorData.Mark = 2;

            SelectData mirrorData2 = selMgr.CreateSelectData();
            mirrorData2.Mark = 1;
            

            // Entities to mirror need mark 1 (can use null for Data if no special marks needed)
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
            selected = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", -halfWidth, topY-Mm(50), 0, false, 0, null, 0);
            swSketchManager.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimClosest, -halfWidth, Mm(325.0), 0);
             //following up 
             // 1. Select the line (using coordinates or name)
            // Select the line
            
            bool status = swModel.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, bottomY, 0, false, 0, null, 0);
            swModel.SketchManager.CreateConstructionGeometry();

            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
            SketchSegment line3 = (SketchSegment)swSketchManager.CreateLine(-halfWidth, Mm(320), 0, 0, 0, 0);
            SketchSegment line4 = (SketchSegment)swSketchManager.CreateLine(halfWidth, Mm(320), 0, 0, 0, 0);
            line3.Select4(false, null);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0,  0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");
            swModel.ClearSelection2(true);

            swSketchManager.InsertSketch(true);
            swModel.Extension.SelectByID2("Sketch2", "SKETCH", 0, 0, 0, false, 0, null, 0);
            swModel.FeatureManager.FeatureCut4(false, false, false, (int)swEndConditions_e.swEndCondThroughAll, (int)swEndConditions_e.swEndCondThroughAll, 0, 0, false, false, false, false, 0, 0, false, false, false, false, false, true, true, true, true, false, 0, 0, false, false);
            swModel.Extension.SelectByID2("Z-Achse", "AXIS", 0, 0, 0, false, 1, null, 0);
            swModel.Extension.SelectByID2("Cut-Extrude1", "BODYFEATURE", 0, 0, 0, true, 4, null, 0);
            Feature myPattern = (Feature)swModel.FeatureManager.FeatureCircularPattern5(
                60,              // 1. Total instances
                2 * Math.PI,     // 2. 360 degrees
                false,           // 3. Reverse direction
                "",              // 4. Instances to skip
                true,            // 5. Equal spacing
                true,            // 6. Verify
                false,           // 7. Geometry pattern (Set to FALSE for first test)
                true,            // 8. Vary sketch
                false,           // 9. Synchronize orientation
                false,           // 10. Direction 2 Reverse
                0,               // 11. Instances (Dir 2)
                0.0,             // 12. Angle (Dir 2)
                "",              // 13. Instances to skip (Dir 2)
                false            // 14. Spacing 2
            );


          


                        // Exit sketch
            swSketchManager.InsertSketch(true);
            
            // ---------------------------------------------------------------------
            // Save the part
            // ---------------------------------------------------------------------
            string fullPath = Path.Combine(outFolder, "StatorBleche.SLDPRT");
            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, 
                           (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
            
            Console.WriteLine($"Part saved to: {fullPath}");
            
            // Re-enable sketch inference
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true);
            
            Console.WriteLine("Done!");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fatal error: " + ex.Message);
            
            // Make sure to re-enable inference even on error
            if (swApp != null)
            {
                try { swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, true); } catch { }
            }
        }
    }
}
