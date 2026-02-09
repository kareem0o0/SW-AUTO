using System;

using System.IO;

using SolidWorks.Interop.sldworks;

using SolidWorks.Interop.swconst;



namespace SwAutomation

{

    class Program

    {

        static void Main(string[] args)

        {

            string outFolder = @"D:\work\birr machines\parts";

            try

            {

                if (!Directory.Exists(outFolder)) Directory.CreateDirectory(outFolder);

                SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));

                swApp.Visible = true;



                // Fixed parameters for Part A

                string partAPath = GeneratePart(swApp, "Part_A", "100", "50", "10", "5", "3", outFolder);



                // Fixed parameters for Part B

                string partBPath = GeneratePart(swApp, "Part_B", "80", "80", "20", "4", "2", outFolder);



                // Create assembly with both parts

                GenerateAssembly(swApp, partAPath, partBPath, outFolder);



                Console.WriteLine($"Success! Parts and Assembly saved to: {outFolder}");

            }

            catch (Exception ex)

            {

                Console.WriteLine("Error: " + ex.Message);

            }

        }



        static string GeneratePart(SldWorks swApp, string name, string sx, string sy, string sz, string sDia, string sCount, string folder)

        {

            if (!double.TryParse(sx, out double x) || !double.TryParse(sy, out double y) || !double.TryParse(sz, out double z)) return "";

            string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            ModelDoc2 swModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);

            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);

            swModel.SketchManager.InsertSketch(true);

            swModel.SketchManager.CreateCornerRectangle(0, 0, 0, x / 1000, y / 1000, 0);

            if (double.TryParse(sDia, out double hDia) && int.TryParse(sCount, out int hCount) && hCount > 0)

            {

                double spacingX = (x / 1000) / (hCount + 1);

                for (int i = 1; i <= hCount; i++)

                    swModel.SketchManager.CreateCircleByRadius(spacingX * i, (y / 2000), 0, (hDia / 2000));

            }

            swModel.FeatureManager.FeatureExtrusion2(true, false, false, (int)swEndConditions_e.swEndCondBlind, 0, z / 1000, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);

            string fullPath = Path.Combine(folder, name + ".SLDPRT");

            swModel.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);

            System.Threading.Thread.Sleep(500); // Wait for file to be written before assembly references it

            return fullPath;

        }



        static void GenerateAssembly(SldWorks swApp, string partAPath, string partBPath, string folder)
{
    string template = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
    ModelDoc2 assyModel = (ModelDoc2)swApp.NewDocument(template, 0, 0, 0);
    AssemblyDoc swAssy = (AssemblyDoc)assyModel;

    // 1. Insert Components and catch the Component object immediately
    // Note: AddComponent4 is more robust than AddComponent
    Component2 compA = swAssy.AddComponent5(partAPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
    Component2 compB = swAssy.AddComponent5(partBPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0.05);

    assyModel.ClearSelection2(true);

    // 2. Select Planes using the documented naming convention: "PlaneName@ComponentName@AssemblyName"
    // We will mate the Front Plane of Part A to the Front Plane of Part B
    string assemblyName = assyModel.GetTitle();
    string partAName = compA.Name2;
    string partBName = compB.Name2;

    // Select Front Plane of Part A
    bool status1 = assyModel.Extension.SelectByID2("Front Plane@" + partAName + "@" + assemblyName, "PLANE", 0, 0, 0, false, 1, null, 0);
    // Select Front Plane of Part B (Append to selection using 'true')
    bool status2 = assyModel.Extension.SelectByID2("Front Plane@" + partBName + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);

    if (status1 && status2)
    {
        int mateError = 0;
        // 3. Apply the Mate (Using AddMate3 as per the documentation logic)
        swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, 
                       (int)swMateAlign_e.swMateAlignALIGNED, 
                       false, 0, 0, 0, 0, 0, 0, 0, 0, false, out mateError);

        if (mateError == 0) Console.WriteLine("Mate 1 (Front Planes) Success!");
    }

    // Repeat for Top Plane to fully constrain
    assyModel.ClearSelection2(true);
    status1 = assyModel.Extension.SelectByID2("Top Plane@" + partAName + "@" + assemblyName, "PLANE", 0, 0, 0, false, 1, null, 0);
    status2 = assyModel.Extension.SelectByID2("Top Plane@" + partBName + "@" + assemblyName, "PLANE", 0, 0, 0, true, 1, null, 0);
    
    if (status1 && status2)
    {
        int mateError = 0;
        swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out mateError);
    }

    assyModel.ForceRebuild3(false);
    string assemblyPath = Path.Combine(folder, "Final_Assembly.SLDASM");
    assyModel.SaveAs3(assemblyPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);
}

    }

}