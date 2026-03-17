using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

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
internal sealed class AutomationUiScope : IDisposable
{
    private readonly SldWorks _swApp;
    private readonly bool _originalCommandInProgress;
    private readonly bool _originalInputDimValOnCreate;
    private readonly bool _originalEnableConfirmationCorner;
    private readonly bool _originalSketchPreviewDimensionOnSelect;

    public AutomationUiScope(SldWorks swApp)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));

        // Save the current SolidWorks UI state so we can restore it later exactly as it was.
        _originalCommandInProgress = _swApp.CommandInProgress;
        _originalInputDimValOnCreate = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate);
        _originalEnableConfirmationCorner = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner);
        _originalSketchPreviewDimensionOnSelect = _swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect);

        // Turn off the interactive options that commonly interfere with automation.
        _swApp.CommandInProgress = true;
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner, false);
        _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect, false);
    }

    public void Dispose()
    {
        try
        {
            // Restore the user's original settings even if part creation failed.
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, _originalInputDimValOnCreate);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swEnableConfirmationCorner, _originalEnableConfirmationCorner);
            _swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swSketchPreviewDimensionOnSelect, _originalSketchPreviewDimensionOnSelect);
        }
        finally
        {
            // Mark the command scope as finished last.
            _swApp.CommandInProgress = _originalCommandInProgress;
        }
    }
}

