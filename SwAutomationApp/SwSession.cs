// Import base .NET types; alternative: remove if unused to speed build slightly.
using System;
// Import COM cleanup helper for deterministic release of COM wrappers.
using System.Runtime.InteropServices;
// Import SOLIDWORKS main COM interop types.
using SolidWorks.Interop.sldworks;

namespace SwAutomation;

// Wraps SOLIDWORKS COM application lifecycle (connect, visibility, cleanup).
public sealed class SwSession : IDisposable
{
    // Expose connected SOLIDWORKS app to builders.
    public SldWorks Application { get; }

    // Tracks whether Dispose already ran.
    private bool _disposed;

    // Connect to SOLIDWORKS by ProgID and configure visibility.
    public SwSession(bool visible)
    {
        try
        {
            // Resolve SOLIDWORKS COM type.
            Type swType = Type.GetTypeFromProgID("SldWorks.Application")
                ?? throw new InvalidOperationException("SOLIDWORKS ProgID was not found.");

            // Create (or connect to) SOLIDWORKS application object.
            object swObject = Activator.CreateInstance(swType)
                ?? throw new InvalidOperationException("Failed to create SOLIDWORKS COM instance.");
            Application = (SldWorks)swObject;

            // Apply requested UI visibility.
            Application.Visible = visible;
        }
        catch (Exception ex)
        {
            // Raise a critical startup error to Program.Main.
            throw new InvalidOperationException("Failed to connect to SOLIDWORKS.", ex);
        }
    }

    // Dispose COM resources.
    public void Dispose()
    {
        // Prevent duplicate cleanup calls.
        if (_disposed) return;

        try
        {
            // Release COM wrapper when available.
            if (Application != null)
            {
                Marshal.FinalReleaseComObject(Application);
            }
        }
        catch (Exception ex)
        {
            // Cleanup errors are non-fatal and logged locally.
            Console.WriteLine("SOLIDWORKS cleanup warning: " + ex.Message);
        }

        // Mark object as cleaned up.
        _disposed = true;
    }
}
