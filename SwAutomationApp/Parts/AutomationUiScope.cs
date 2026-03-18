// Gives access to IDisposable, which lets this helper clean itself up automatically
// when it leaves a using block.
using System;
// Gives access to the main SolidWorks COM application type, SldWorks.
using SolidWorks.Interop.sldworks;
// Gives access to the SolidWorks enum values used for UI preference toggles.
using SolidWorks.Interop.swconst;

// All automation files use the same namespace so they can reference each other directly.
namespace SwAutomation;

/// <summary>
/// Temporarily disables some interactive SolidWorks UI behavior while automation is running.
///
/// Why this matters:
/// SolidWorks is normally designed for a human user clicking through dialogs and previews.
/// During automation, those interactive behaviors can interrupt scripts or make geometry creation
/// unstable. This helper creates a safe "quiet mode" around part generation.
///
/// The constructor stores the user's original UI settings.
/// Dispose() restores them after automation finishes.
/// </summary>
// internal = used only inside this project.
// sealed = no inheritance is needed here.
// IDisposable = lets us use "using var ..." so restoration happens automatically.
internal sealed class AutomationUiScope : IDisposable
{
    // Reference to the live SolidWorks application that this helper will configure.
    private readonly SldWorks _swApp;
    // Stores whether SolidWorks already thought a command was running before automation started.
    private readonly bool _originalCommandInProgress;
    // Stores the original "input dimension while creating" setting.
    private readonly bool _originalInputDimValOnCreate;
    // Stores the original confirmation-corner setting.
    private readonly bool _originalEnableConfirmationCorner;
    // Stores the original sketch preview dimension setting.
    private readonly bool _originalSketchPreviewDimensionOnSelect;
    // Stores the original sketch inference setting.
    private readonly bool _originalSketchInference;

    // The constructor runs when automation enters a quiet UI scope.
    public AutomationUiScope(SldWorks swApp)
    {
        // Save the SolidWorks application reference so the rest of this helper can use it.
        _swApp = swApp;

        // Read and store the current SolidWorks UI state so we can restore it later exactly as it was.
        // CommandInProgress tells SolidWorks that automation is actively controlling the session.
        _originalCommandInProgress = _swApp.CommandInProgress;
        // This setting controls whether SolidWorks asks for dimensions interactively while sketching.
        _originalInputDimValOnCreate = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
        // This setting controls the confirmation corner that normally appears during manual editing.
        _originalEnableConfirmationCorner = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner);
        // This setting controls the preview dimensions shown while selecting sketch entities.
        _originalSketchPreviewDimensionOnSelect = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect);
        // This setting controls automatic sketch inference such as auto-snaps and auto-relations.
        _originalSketchInference = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference);

        // Tell SolidWorks that a controlled automation command is in progress.
        _swApp.CommandInProgress = true;
        // Disable interactive dimension prompts because automation sets dimensions itself.
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
        // Disable the confirmation corner because automation does not click Accept/Cancel manually.
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner, false);
        // Disable preview dimensions so sketching and selection stay quieter and more predictable.
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect, false);
        // Disable sketch inference so SolidWorks does not add unexpected automatic relations.
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, false);
    }

    // Dispose is called automatically at the end of a using block.
    // Its job is to restore the user's original SolidWorks state.
    public void Dispose()
    {
        // Restore the original "input dimension while creating" behavior.
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, _originalInputDimValOnCreate);
        // Restore the original confirmation corner behavior.
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner, _originalEnableConfirmationCorner);
        // Restore the original sketch preview dimension behavior.
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect, _originalSketchPreviewDimensionOnSelect);
        // Restore the original sketch inference behavior.
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchInference, _originalSketchInference);
        // Finally restore the original command-in-progress state.
        _swApp.CommandInProgress = _originalCommandInProgress;
    }
}
