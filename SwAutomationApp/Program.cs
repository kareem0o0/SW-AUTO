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

        SldWorks swApp = new SldWorks();
        swApp.Visible = true;

        PdmModule pdm = new PdmModule();
        string outFolder = @"C:\Users\kareem.salah\Downloads\birr machines\birr machines\parts";

        SkeletonPart skeleton = new SkeletonPart(swApp, pdm);
        skeleton.OutputFolder = outFolder;
        skeleton.DESideOffset = 500.0;
        skeleton.NDESideOffset = 500.0;
        skeleton.GroundOffsetMm = -250.0;
        skeleton.CloseAfterCreate = true;

        AssemblyFile rotor = new AssemblyFile(swApp, pdm);
        rotor.FileName = "RotorComplete.SLDASM";
        rotor.OutputFolder = outFolder;
        rotor.CloseAfterCreate = true;

        AssemblyFile stator = new AssemblyFile(swApp, pdm);
        stator.FileName = "StatorComplete.SLDASM";
        stator.OutputFolder = outFolder;
        stator.CloseAfterCreate = true;

        AssemblyFile housing = new AssemblyFile(swApp, pdm);
        housing.FileName = "HousingMachined.SLDASM";
        housing.OutputFolder = outFolder;
        housing.CloseAfterCreate = true;

        AssemblyFile machine = new AssemblyFile(swApp, pdm);
        machine.FileName = "MachineAssembly.SLDASM";
        machine.OutputFolder = outFolder;

        string skeletonPath = skeleton.Create();
        string rotorPath = rotor.Create();
        string statorPath = stator.Create();
        string housingPath = housing.Create();
        string machinePath = machine.Create();

        var inserted = machine.Insert(skeletonPath);
        machine.MateToOrigin(inserted);

        inserted = machine.Insert(rotorPath);
        machine.MateToOrigin(inserted);

        inserted = machine.Insert(statorPath);
        machine.MateToOrigin(inserted);

        inserted = machine.Insert(housingPath);
        machine.MateToOrigin(inserted);


    }
}
