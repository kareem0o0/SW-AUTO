using System;
using SolidWorks.Interop.sldworks;
using SwAutomation.Pdm;

[assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")]

namespace SwAutomation;

/// <summary>
/// This file is the executable entry point of the whole application.
///
/// It stays intentionally small:
/// 1. create the external services we need
/// 2. choose which macro flow should run
/// 3. hand control to that macro
///
/// Keeping startup simple makes it much easier to switch between test flows later.
/// </summary>
public static class MachineAssembly
{
    [STAThread]
    public static void Main(string[] args)
    {
        // SolidWorks must be created first because every generated object uses this live COM app.
        Console.WriteLine("Connecting to SOLIDWORKS...");

        SldWorks swApp = new SldWorks();
        swApp.Visible = true;

        // PDM is created once here and then shared with all generator objects.
        // That keeps vault logic centralized in one service object.
        PdmModule pdm = new PdmModule();
        string localOutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";

        // Run4 is the current main machine-build flow.
        // Change this one line if you want a different startup macro later.
        Project1.Run4(localOutputFolder, swApp, pdm);
    }
}

