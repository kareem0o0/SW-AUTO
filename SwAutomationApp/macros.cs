using System;
using SolidWorks.Interop.sldworks;
using SwAutomation.Pdm;

namespace SwAutomation;

public static class Project1
{
    public static void Run(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        var skeleton = new SkeletonPart(swApp, pdm)
        {
            OutputFolder = outFolder,
            SideOffsetMm = 500.0,
            GroundOffsetMm = -250.0,
            CloseAfterCreate = true
        };
        var rotor = new AssemblyDocumentDefinition(swApp, pdm, "RotorComplete.SLDASM")
        {
            OutputFolder = outFolder,
            CloseAfterCreate = true
        };
        var stator = new AssemblyDocumentDefinition(swApp, pdm, "StatorComplete.SLDASM")
        {
            OutputFolder = outFolder,
            CloseAfterCreate = true
        };
        var housing = new AssemblyDocumentDefinition(swApp, pdm, "HousingMachined.SLDASM")
        {
            OutputFolder = outFolder,
            CloseAfterCreate = true
        };
        var machine = new AssemblyDocumentDefinition(swApp, pdm, "MachineAssembly.SLDASM")
        {
            OutputFolder = outFolder
        };

        string skeletonPath = skeleton.Create();
        string rotorPath = rotor.Create();
        string statorPath = stator.Create();
        string housingPath = housing.Create();
        string machinePath = machine.Create();

        var inserted = machine.InsertComponentToOpenAssembly(skeletonPath);
        machine.MatePlanes(inserted);

        inserted = machine.InsertComponentToOpenAssembly(rotorPath);
        machine.MatePlanes(inserted);

        inserted = machine.InsertComponentToOpenAssembly(statorPath);
        machine.MatePlanes(inserted);

        inserted = machine.InsertComponentToOpenAssembly(housingPath);
        machine.MatePlanes(inserted);

        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }

    public static void Run2(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        var skeleton = new SkeletonPart(swApp, pdm)
        {
            OutputFolder = outFolder,
            SideOffsetMm = 2000.0,
            GroundOffsetMm = -500.0,
            CloseAfterCreate = true
        };
        var statorSheet = new StatorSheetPart(swApp, pdm)
        {
            OutputFolder = outFolder,
            CloseAfterCreate = true
        };
        var shaft = new ShaftPart(swApp, pdm)
        {
            OutputFolder = outFolder,
            CloseAfterCreate = true
        };
        var machine = new AssemblyDocumentDefinition(swApp, pdm, "MachineAssembly.SLDASM")
        {
            OutputFolder = outFolder
        };

        string skeletonPath = skeleton.Create();
        string statorSheetPath = statorSheet.Create();
        string shaftPath = shaft.Create();
        string machinePath = machine.Create();

        var insertedSkeleton = machine.InsertComponentToOpenAssembly(skeletonPath);
        machine.MatePlanes(insertedSkeleton);
        var insertedStatorSheet = machine.InsertComponentToOpenAssembly(statorSheetPath);
        var insertedShaft = machine.InsertComponentToOpenAssembly(shaftPath);
        machine.ApplyCoincidentMate(insertedSkeleton, "X-Achse", insertedStatorSheet, "Z-Achse");
        machine.ApplyCoincidentMate(insertedSkeleton, "Ebene rechts", insertedStatorSheet, "Ebene vorne");
        machine.ApplyCoincidentMate(insertedSkeleton, "Ebene vorne", insertedStatorSheet, "Ebene rechts");
        machine.ApplyCoincidentMate(insertedSkeleton, "X-Achse", insertedShaft, "Z-Achse");

        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }

    public static void Run3(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        _ = swApp;
        _ = pdm;
        Console.WriteLine("Run3 is reserved for ad-hoc generator checks.");
    }

    public static void Run4(string outFolder, SldWorks swApp, PdmModule pdm)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        var machineAssembly = new MachineAssemblyDefinition(swApp, pdm)
        {
            OutputFolder = outFolder
        };

        string machinePath = machineAssembly.Create();
        Console.WriteLine($"Macro completed. Machine assembly: {machinePath}");
    }
}
