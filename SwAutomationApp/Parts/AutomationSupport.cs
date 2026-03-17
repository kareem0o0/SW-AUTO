using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation;

/// <summary>
/// Small shared helper methods used by many part and assembly classes.
///
/// These helpers stay tiny on purpose.
/// The goal is to avoid repeating the same low-level checks
/// in every generator file.
/// </summary>
internal static class AutomationSupport
{
    /// <summary>
    /// Ensures that required text settings such as output folders or file names are present.
    ///
    /// We use one shared method so all generators produce clear and consistent error messages.
    /// </summary>
    public static string RequireText(string value, string propertyName, string ownerName)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new InvalidOperationException($"{ownerName} requires a non-empty {propertyName} value.");

        return value;
    }
}

