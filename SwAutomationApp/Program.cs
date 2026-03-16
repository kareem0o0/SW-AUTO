using System;
using SolidWorks.Interop.sldworks;
using SwAutomation.Pdm;

[assembly: System.Runtime.Versioning.SupportedOSPlatform("windows")]

namespace SwAutomation;

public static class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        Console.WriteLine("Connecting to SOLIDWORKS...");

        SldWorks swApp = new SldWorks
        {
            Visible = true
        };

        var pdm = new PdmModule();
        string localOutputFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";

        Project1.Run4(localOutputFolder, swApp, pdm);
    }
}
