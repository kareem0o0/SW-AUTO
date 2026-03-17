using SolidWorks.Interop.sldworks;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Named wrapper for the housing assembly.
///
/// This keeps the structure clear today and gives us a dedicated location
/// for housing-specific behavior in the future.
/// </summary>
public sealed class HousingAssembly : AssemblyFile
{
    public HousingAssembly(SldWorks swApp, PdmModule pdm) : base(swApp, pdm)
    {
        // Set the default document file name for this assembly type.
        FileName = "HousingMachined.SLDASM";
    }
}

