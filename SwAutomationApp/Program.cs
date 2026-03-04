// Import base .NET types; alternative: remove if unused to speed build slightly.
using System;
using SolidWorks.Interop.sldworks;
// Removes platform warnings by specifying this code is Windows-only.
[assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")]

namespace SwAutomation;

// Program entry type; coordinates session and service classes.
public static class Program
{
    [STAThread]
    
    public static void Main(string[] args)
{
    Console.WriteLine("Connecting to SOLIDWORKS...");
    SldWorks swApp = new SldWorks { Visible = true };

    var pdm = new global::SwAutomation.Pdm.PdmModule();
    var myPart = new Part(swApp, pdm);
    var assembly = new Assembly(swApp, pdm);

    pdm.Login(); 
    string outFolder = @"60_Tests\665_Test_Kareem";

    //Project1.Run(outFolder, myPart, assembly);
    //myPart.Create_stator_sheet(outFolder);
    //pdm.GetDataCardValues(@"60_Tests\665_Test_Kareem\BMZS010258.sldprt");
    string statorFileName = myPart.Create_stator_sheet(outFolder,true);
    pdm.FillBirrDataCard(System.IO.Path.Combine(outFolder, statorFileName));

    //myPart.CreateSkeleton(sideOffset: 500.0, groundOffset: -250.0, outFolder: outFolder,closeAfterCreate: true);
    //string myLocalPart = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts\StatorComplete.SLDASM";
    //pdm.AddExistingFileToPdm(myLocalPart, @"60_Tests\665_Test_Kareem");
}
}
