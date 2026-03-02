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
        string outFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";
        Console.WriteLine("Connecting to SOLIDWORKS...");
        SldWorks swApp = new SldWorks();
        swApp.Visible = true;

        var parts = new Parts(swApp);

        parts.creat_stator_sheet(outFolder);
        parts.CreateReferencePlanes(sideOffset: 500.0, groundOffset: -250.0, outFolder: outFolder); // values in mm
        
        
    }
}
