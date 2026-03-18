using System;
using System.Collections.Generic;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Static drawing helper methods.
///
/// The current drawing implementation is intentionally kept outside the part files.
/// This keeps the part classes focused on owning data, while this file acts as the
/// drawing engine that reads that data and builds a drawing document.
/// </summary>
internal static class DrawingMethods
{
    /// <summary>
    /// Creates the torsion-bar drawing from the values stored on a TorsionBarPart object.
    ///
    /// High-level flow:
    /// 1. make sure the part exists
    /// 2. choose paper size and scale
    /// 3. resolve drawing template and sheet format
    /// 4. create the drawing
    /// 5. insert the required views
    /// 6. save locally or to PDM
    /// </summary>
    public static string CreateTorsionBarDrawing(TorsionBarPart part, SldWorks swApp, PdmModule pdm)
    {
        // The drawing may have its own output folder, but if it is blank we reuse the part folder.
        string outFolder = string.IsNullOrWhiteSpace(part.DrawingOutputFolder)
            ? part.OutputFolder
            : Path.GetFullPath(part.DrawingOutputFolder);
        Directory.CreateDirectory(outFolder);

        // If a custom drawing file name was not given, reuse the part file name with a drawing extension.
        string drawingFileName = string.IsNullOrWhiteSpace(part.DrawingLocalFileName)
            ? Path.ChangeExtension(part.LocalFileName, ".SLDDRW")
            : part.DrawingLocalFileName;

        // Always create or refresh the part first so the drawing references the latest geometry.
        string partFileName = part.CreatePart();
        if (string.IsNullOrWhiteSpace(partFileName))
            throw new InvalidOperationException("CreateTorsionBarDrawing could not create the source TorsionBarPart.");

        string partPath = Path.GetFullPath(Path.Combine(part.OutputFolder, partFileName));
        if (!File.Exists(partPath))
            throw new FileNotFoundException("The source Torsion Bar part file was not found after creation.", partPath);

        // Keep the source part loaded in SolidWorks while the drawing stays open.
        // This makes projected and manual view operations more reliable.
        ModelDoc2 sourcePartModel = swApp.IGetOpenDocumentByName2(partPath);
        if (sourcePartModel == null)
        {
            int openErrors = 0;
            int openWarnings = 0;
            sourcePartModel = swApp.OpenDoc6(
                partPath,
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                string.Empty,
                ref openErrors,
                ref openWarnings);

            if (sourcePartModel == null)
                throw new InvalidOperationException($"Could not keep the torsion-bar part open for drawing work. OpenDoc6 error code: {openErrors}.");
        }

        string languageCode = (part.DrawingLanguageCode ?? string.Empty).Trim().ToUpperInvariant();
        if (languageCode != "EN" && languageCode != "DE")
            throw new InvalidOperationException("TorsionBarPart.DrawingLanguageCode must be either EN or DE.");

        // Candidate sheet sizes and their dimensions in meters.
        // The code walks from small to large so it picks the smallest sheet that still fits.
        (string PaperCode, swDwgPaperSizes_e PaperSize, double Width, double Height)[] sheets =
        {
            ("A4", swDwgPaperSizes_e.swDwgPaperA4size, 0.21, 0.297),
            ("A3", swDwgPaperSizes_e.swDwgPaperA3size, 0.297, 0.42),
            ("A2", swDwgPaperSizes_e.swDwgPaperA2size, 0.42, 0.594),
            ("A1", swDwgPaperSizes_e.swDwgPaperA1size, 0.594, 0.841)
        };
        double[] scaleDenominators = { 1.0, 2.0, 3.0, 4.0, 5.0, 10.0 };
        const double sideMargin = 0.02;
        const double leftMargin = 0.025;
        const double topMargin = 0.045;
        const double projectedViewGap = 0.012;
        double bottomTitleBlockClearance = Math.Max(0.0, part.DrawingBottomTitleBlockClearance);

        string paperCode = string.Empty;
        swDwgPaperSizes_e paperSize = swDwgPaperSizes_e.swDwgPaperA3size;
        double scaleNumerator = 1.0;
        double scaleDenominator = 1.0;
        double frontX = 0.0;
        double frontY = 0.0;
        double topX = 0.0;
        double topY = 0.0;
        double sideX = 0.0;
        double sideY = 0.0;
        bool layoutFound = false;

        // Try progressively larger sheets and a list of scale options until the three views fit.
        // Once one combination works, the loop stops and uses that layout.
        foreach (var sheet in sheets)
        {
            foreach (double candidateScaleDenominator in scaleDenominators)
            {
                double scale = 1.0 / candidateScaleDenominator;
                double frontWidth = part.BarLength * scale;
                double frontHeight = part.BarHeight * scale;
                double topHeight = part.BarThickness * scale;
                double sideWidth = part.BarThickness * scale;
                double totalWidth = leftMargin + frontWidth + projectedViewGap + sideWidth + sideMargin;
                double totalHeight = bottomTitleBlockClearance + frontHeight + projectedViewGap + topHeight + topMargin;

                if (totalWidth > sheet.Width || totalHeight > sheet.Height)
                    continue;

                paperCode = sheet.PaperCode;
                paperSize = sheet.PaperSize;
                scaleDenominator = candidateScaleDenominator;

                if (part.DrawingUseFirstAngleProjection)
                {
                    frontX = leftMargin + projectedViewGap + sideWidth + (frontWidth / 2.0);
                    sideX = leftMargin + (sideWidth / 2.0);
                    frontY = sheet.Height - topMargin - (frontHeight / 2.0);
                    topY = frontY - (frontHeight / 2.0) - projectedViewGap - (topHeight / 2.0);
                }
                else
                {
                    frontX = leftMargin + (frontWidth / 2.0);
                    sideX = frontX + (frontWidth / 2.0) + projectedViewGap + (sideWidth / 2.0);
                    topY = sheet.Height - topMargin - (topHeight / 2.0);
                    frontY = topY - (topHeight / 2.0) - projectedViewGap - (frontHeight / 2.0);
                }

                topX = frontX;
                sideY = frontY;
                layoutFound = true;
                break;
            }

            if (layoutFound)
                break;
        }

        if (!layoutFound)
            throw new InvalidOperationException("Could not find a drawing sheet size and scale that fits the Torsion Bar front, top, and side views.");

        // Start with SolidWorks' default drawing template if one is configured.
        string defaultDrawingTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing) ?? string.Empty;

        string drawingTemplatePath = string.Empty;
        if (!string.IsNullOrWhiteSpace(part.DrawingTemplatePathOverride))
        {
            // Use the manual override when the caller wants a specific template file.
            drawingTemplatePath = Path.GetFullPath(part.DrawingTemplatePathOverride);
            if (!File.Exists(drawingTemplatePath))
                throw new FileNotFoundException("The drawing template override file was not found.", drawingTemplatePath);
        }
        else if (!string.IsNullOrWhiteSpace(defaultDrawingTemplate) && File.Exists(defaultDrawingTemplate))
        {
            drawingTemplatePath = defaultDrawingTemplate;
        }
        else
        {
            // Otherwise search the configured SolidWorks document-template folders.
            string documentTemplateFolders = swApp.GetUserPreferenceStringListValue((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates) ?? string.Empty;

            char[] separators = { '|', ';', '\r', '\n' };
            foreach (string rawFolder in documentTemplateFolders.Split(separators, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (!Directory.Exists(rawFolder))
                    continue;

                string[] templates = Directory.GetFiles(rawFolder, "*.drwdot", SearchOption.TopDirectoryOnly);
                if (templates.Length == 0)
                    continue;

                drawingTemplatePath = templates[0];
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(drawingTemplatePath))
            throw new InvalidOperationException("Could not resolve a SolidWorks drawing document template. Set DrawingTemplatePathOverride or configure a default drawing template in SolidWorks.");

        string sheetFormatPath = string.Empty;
        if (!string.IsNullOrWhiteSpace(part.DrawingSheetFormatPathOverride))
        {
            // Direct override for the sheet format file.
            sheetFormatPath = Path.GetFullPath(part.DrawingSheetFormatPathOverride);
            if (!File.Exists(sheetFormatPath))
                throw new FileNotFoundException("The sheet-format override file was not found.", sheetFormatPath);
        }
        else
        {
            // Build a list of folders where the expected sheet format might exist.
            // SolidWorks folders are checked first, then the Birr template folder set on the object.
            List<string> sheetFormatFolders = new();
            char[] separators = { '|', ';', '\r', '\n' };

            if (part.DrawingPreferSolidWorksTemplateLocations)
            {
                string configuredSheetFormatFolders = swApp.GetUserPreferenceStringListValue((int)swUserPreferenceStringValue_e.swFileLocationsSheetFormat) ?? string.Empty;
                string configuredNewSheetFormatFolders = swApp.GetUserPreferenceStringListValue((int)swUserPreferenceStringValue_e.swFileLocationsNewSheetFormat) ?? string.Empty;

                foreach (string rawFolder in configuredSheetFormatFolders.Split(separators, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                {
                    if (Directory.Exists(rawFolder))
                        sheetFormatFolders.Add(Path.GetFullPath(rawFolder));
                }

                foreach (string rawFolder in configuredNewSheetFormatFolders.Split(separators, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                {
                    if (Directory.Exists(rawFolder))
                        sheetFormatFolders.Add(Path.GetFullPath(rawFolder));
                }

                if (!string.IsNullOrWhiteSpace(defaultDrawingTemplate) && File.Exists(defaultDrawingTemplate))
                {
                    string templateFolder = Path.GetDirectoryName(defaultDrawingTemplate) ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(templateFolder) && Directory.Exists(templateFolder))
                        sheetFormatFolders.Add(Path.GetFullPath(templateFolder));
                }
            }

            if (!string.IsNullOrWhiteSpace(part.DrawingTemplateFolderPath) && Directory.Exists(part.DrawingTemplateFolderPath))
                sheetFormatFolders.Add(Path.GetFullPath(part.DrawingTemplateFolderPath));

            string expectedSheetFormatFileName = $"{paperCode}_Birr_Machines_{languageCode}.slddrt";
            foreach (string folder in sheetFormatFolders)
            {
                // Check both direct and recursive matches because company template folders can vary.
                string directPath = Path.Combine(folder, expectedSheetFormatFileName);
                if (File.Exists(directPath))
                {
                    sheetFormatPath = directPath;
                    break;
                }

                string[] recursiveMatches = Directory.GetFiles(folder, expectedSheetFormatFileName, SearchOption.AllDirectories);
                if (recursiveMatches.Length > 0)
                {
                    sheetFormatPath = recursiveMatches[0];
                    break;
                }
            }

            if (string.IsNullOrWhiteSpace(sheetFormatPath))
                throw new FileNotFoundException($"Could not find sheet format '{expectedSheetFormatFileName}'. Check DrawingLanguageCode or DrawingTemplateFolderPath.");
        }

        ModelDoc2 drawingModel = (ModelDoc2)swApp.NewDocument(drawingTemplatePath, 0, 0, 0);
        if (drawingModel == null)
            throw new InvalidOperationException("Failed to create a new drawing document.");

        DrawingDoc drawingDoc = drawingModel as DrawingDoc;
        if (drawingDoc == null)
        {
            swApp.CloseDoc(drawingModel.GetTitle());
            throw new InvalidOperationException("Could not access the SolidWorks drawing document.");
        }

        string sheetName = string.IsNullOrWhiteSpace(part.DrawingSheetName) ? "Sheet1" : part.DrawingSheetName;

        // Apply the chosen sheet size, scale, projection method, and company sheet format.
        bool sheetConfigured = drawingDoc.SetupSheet6(
            sheetName,
            (int)paperSize,
            (int)swDwgTemplates_e.swDwgTemplateCustom,
            scaleNumerator,
            scaleDenominator,
            part.DrawingUseFirstAngleProjection,
            sheetFormatPath,
            0,
            0,
            string.Empty,
            false,
            0,
            0,
            0,
            0,
            1,
            1);

        if (!sheetConfigured)
            throw new InvalidOperationException("Could not configure the drawing sheet.");

        Console.WriteLine($"Drawing template: {drawingTemplatePath}");
        Console.WriteLine($"Sheet format: {sheetFormatPath}");

        // Activate the sheet before inserting views so every view lands on the expected sheet.
        drawingDoc.ActivateSheet(sheetName);
        drawingDoc.EditSheet();
        drawingModel.EditRebuild3();

        // Insert the main front view first because the projected views depend on it.
        View frontView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Front", frontX, frontY, 0);
        if (frontView == null)
            throw new InvalidOperationException("Could not create the front torsion-bar view.");

        if (!string.IsNullOrWhiteSpace(part.DrawingReferencedConfiguration))
            frontView.ReferencedConfiguration = part.DrawingReferencedConfiguration;

        frontView.Position = new double[] { frontX, frontY, 0.0 };
        frontView.PositionLocked = false;

        // SolidWorks creates projected views from the currently selected base view.
        drawingModel.ClearSelection2(true);
        if (!drawingDoc.ActivateView(frontView.Name))
            throw new InvalidOperationException("Could not activate the front torsion-bar view before creating projected views.");
        if (!drawingModel.Extension.SelectByID2(frontView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0))
            throw new InvalidOperationException("Could not select the front torsion-bar view before creating the projected top view.");

        View topView = drawingDoc.CreateUnfoldedViewAt3(topX, topY, 0, false);
        if (topView == null)
            throw new InvalidOperationException("Could not create the projected top torsion-bar view.");

        drawingModel.ClearSelection2(true);
        if (!drawingDoc.ActivateView(frontView.Name))
            throw new InvalidOperationException("Could not re-activate the front torsion-bar view before creating the side projection.");
        if (!drawingModel.Extension.SelectByID2(frontView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0))
            throw new InvalidOperationException("Could not select the front torsion-bar view before creating the projected side view.");

        View sideView = drawingDoc.CreateUnfoldedViewAt3(sideX, sideY, 0, false);
        if (sideView == null)
            throw new InvalidOperationException("Could not create the projected side torsion-bar view.");

        topView.PositionLocked = false;
        sideView.PositionLocked = false;

        // Rebuild once more so the saved file contains the final view state.
        drawingModel.EditRebuild3();
        drawingDoc.ActivateSheet(sheetName);
        drawingModel.ClearSelection2(true);
        drawingModel.ViewZoomtofit2();

        string savedPath;
        if (part.DrawingSaveToPdm)
        {
            // Drawings can be saved to PDM separately from the part file.
            savedPath = pdm.SaveAsPdm(drawingModel, outFolder, part.DrawingPdmDataCard);
            Console.WriteLine($"Drawing saved to PDM: {savedPath}");
        }
        else
        {
            savedPath = Path.Combine(outFolder, drawingFileName);
            drawingModel.SaveAs3(savedPath, 0, 1);
            Console.WriteLine($"Drawing saved locally: {savedPath}");
        }

        if (part.DrawingCloseAfterCreate)
        {
            swApp.CloseDoc(drawingModel.GetTitle());
            Console.WriteLine("Drawing closed after creating.");
        }

        return Path.GetFileName(savedPath);
    }
}


