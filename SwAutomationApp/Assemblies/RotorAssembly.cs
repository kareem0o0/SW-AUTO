using SolidWorks.Interop.sldworks;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Named wrapper for the rotor assembly.
///
/// Right now this class simply inherits the generic AssemblyFile behavior
/// and provides the default file name for the rotor assembly document.
/// The benefit of keeping it as its own file is that rotor-specific logic can be added later
/// without mixing it into the generic assembly helper.
/// </summary>
public sealed class RotorAssembly : AssemblyFile
{
    public RotorAssembly(SldWorks swApp, PdmModule pdm) : base(swApp, pdm)
    {
        // Set the default document file name for this assembly type.
        FileName = "RotorComplete.SLDASM";
    }
}

