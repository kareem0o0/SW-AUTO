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
        Component2 swComponent = null;
        SldWorks swApp = new SldWorks();
        swApp.Visible = true;

        var part = new Part(swApp);
        var assembly = new Assembly(swApp);

        part.Create_stator_sheet(outFolder);
       /* string skeleton = part.CreateSkeleton(sideOffset: 500.0, groundOffset: -250.0, outFolder: outFolder,closeAfterCreate: true); // values in mm
       /* assembly.CreateAssembly(outFolder, "RotorComplete.SLDASM", closeAfterCreate: true);
        assembly.CreateAssembly(outFolder, "StatorComplete.SLDASM", closeAfterCreate: true);
        assembly.CreateAssembly(outFolder, "HousingMachined.SLDASM", closeAfterCreate: true);
        assembly.CreateAssembly(outFolder, "MachineAssembly.SLDASM", closeAfterCreate: false);
        swComponent = assembly.InsertComponentToOpenAssembly(skeleton);
        assembly.mate_plans(swComponent);
        swComponent = assembly.InsertComponentToOpenAssembly("RotorComplete.SLDASM");
        assembly.mate_plans(swComponent);
        swComponent = assembly.InsertComponentToOpenAssembly("StatorComplete.SLDASM");
        assembly.mate_plans(swComponent);
        swComponent = assembly.InsertComponentToOpenAssembly("HousingMachined.SLDASM");
        assembly.mate_plans(swComponent);*/
        
    }
}
