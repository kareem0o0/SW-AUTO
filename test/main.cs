using System.Runtime.Versioning;
using SolidWorks.Interop.sldworks;

[assembly: SupportedOSPlatform("windows")]

namespace SwAutomation.TestHarness;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Console.WriteLine("Connecting to SOLIDWORKS...");

        SldWorks swApp = new SldWorks
        {
            Visible = true
        };

        Part part = new(swApp);
        Assembly assembly = new(swApp);
        string outFolder = Path.Combine(AppContext.BaseDirectory, "output");

        Macro.Run(outFolder, part, assembly);
    }
}

internal static class Macro
{
    public static void Run(string outFolder, Part part, Assembly assembly)
    {
        if (string.IsNullOrWhiteSpace(outFolder))
            throw new ArgumentException("Output folder is required.", nameof(outFolder));

        string skeleton = part.CreateSkeleton(
            sideOffset: 500.0,
            groundOffset: -250.0,
            outFolder: outFolder,
            closeAfterCreate: true);

        string rotor = assembly.CreateAssembly(outFolder, "RotorComplete.SLDASM", closeAfterCreate: true);
        string stator = assembly.CreateAssembly(outFolder, "StatorComplete.SLDASM", closeAfterCreate: true);
        string housing = assembly.CreateAssembly(outFolder, "HousingMachined.SLDASM", closeAfterCreate: true);
        string machine = assembly.CreateAssembly(outFolder, "MachineAssembly.SLDASM", closeAfterCreate: false);

        Component2 inserted = assembly.InsertComponentToOpenAssembly(skeleton);
        assembly.MatePlans(inserted);

        inserted = assembly.InsertComponentToOpenAssembly(rotor);
        assembly.MatePlans(inserted);

        inserted = assembly.InsertComponentToOpenAssembly(stator);
        assembly.MatePlans(inserted);

        inserted = assembly.InsertComponentToOpenAssembly(housing);
        assembly.MatePlans(inserted);

        Console.WriteLine($"Macro completed. Machine assembly: {machine}");
    }
}
