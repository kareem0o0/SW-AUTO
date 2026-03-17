using SolidWorks.Interop.sldworks;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Named wrapper for the stator assembly.
///
/// It currently reuses the generic AssemblyFile behavior and only applies
/// the default file identity for the stator assembly.
/// </summary>
public sealed class StatorAssembly : AssemblyFile
{
    public StatorAssembly(SldWorks swApp, PdmModule pdm) : base(swApp, pdm)
    {
        // Set the default document file name for this assembly type.
        FileName = "StatorComplete.SLDASM";
    }
}

