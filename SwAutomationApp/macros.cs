using SolidWorks.Interop.sldworks;

namespace SwAutomation;

public static class Project1
{
    public static void Run(string outFolder, Part part, Assembly assembly)
    {
        string skeleton = part.CreateSkeleton(
            sideOffset: 500.0,
            groundOffset: -250.0,
            outFolder: outFolder,
            closeAfterCreate: true);

        assembly.CreateAssembly(outFolder, "RotorComplete.SLDASM", closeAfterCreate: true);
        assembly.CreateAssembly(outFolder, "StatorComplete.SLDASM", closeAfterCreate: true);
        assembly.CreateAssembly(outFolder, "HousingMachined.SLDASM", closeAfterCreate: true);
        assembly.CreateAssembly(outFolder, "MachineAssembly.SLDASM", closeAfterCreate: false);

        Component2 inserted = assembly.InsertComponentToOpenAssembly(skeleton);
        assembly.mate_plans(inserted);

        inserted = assembly.InsertComponentToOpenAssembly("RotorComplete.SLDASM");
        assembly.mate_plans(inserted);

        inserted = assembly.InsertComponentToOpenAssembly("StatorComplete.SLDASM");
        assembly.mate_plans(inserted);

        inserted = assembly.InsertComponentToOpenAssembly("HousingMachined.SLDASM");
        assembly.mate_plans(inserted);
    }
}
