using System;
using System.Collections.Generic;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

internal static class DrawingMethods
{
    public static string CreateTorsionBarDrawing(TorsionBarPart part, SldWorks swApp, PdmModule pdm)
    {
        if (part == null)
            throw new InvalidOperationException("CreateTorsionBarDrawing requires a TorsionBarPart instance.");
        if (swApp == null)
            throw new ArgumentNullException(nameof(swApp));
        if (pdm == null)
            throw new ArgumentNullException(nameof(pdm));

        const double mmToMeters = AutomationSupport.MmToMeters;
        double Mm(double mm) => mm * mmToMeters;

        string outFolder = string.IsNullOrWhiteSpace(part.DrawingOutputFolder)
            ? AutomationSupport.RequireText(part.OutputFolder, nameof(part.OutputFolder), nameof(TorsionBarPart))
            : Path.GetFullPath(part.DrawingOutputFolder);
        Directory.CreateDirectory(outFolder);

        string drawingFileName = string.IsNullOrWhiteSpace(part.DrawingLocalFileName)
            ? Path.ChangeExtension(part.LocalFileName, ".SLDDRW")
            : part.DrawingLocalFileName;

        string partFileName = part.CreatePart();
        if (string.IsNullOrWhiteSpace(partFileName))
            throw new InvalidOperationException("CreateTorsionBarDrawing could not create the source TorsionBarPart.");

        string partPath = Path.GetFullPath(Path.Combine(part.OutputFolder, partFileName));
        if (!File.Exists(partPath))
            throw new FileNotFoundException("The source Torsion Bar part file was not found after creation.", partPath);

        // Keep the source part loaded in SolidWorks while the drawing stays open.
        // Manual projected-view creation is more reliable when the referenced model is already loaded.
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

        (string PaperCode, swDwgPaperSizes_e PaperSize, double WidthMm, double HeightMm)[] sheets =
        {
            ("A4", swDwgPaperSizes_e.swDwgPaperA4size, 210.0, 297.0),
            ("A3", swDwgPaperSizes_e.swDwgPaperA3size, 297.0, 420.0),
            ("A2", swDwgPaperSizes_e.swDwgPaperA2size, 420.0, 594.0),
            ("A1", swDwgPaperSizes_e.swDwgPaperA1size, 594.0, 841.0)
        };
        double[] scaleDenominators = { 1.0, 2.0, 3.0, 4.0, 5.0, 10.0 };
        const double sideMarginMm = 20.0;
        const double leftMarginMm = 25.0;
        const double topMarginMm = 45.0;
        double bottomTitleBlockClearanceMm = Math.Max(0.0, part.DrawingBottomTitleBlockClearanceMm);

        string paperCode = string.Empty;
        swDwgPaperSizes_e paperSize = swDwgPaperSizes_e.swDwgPaperA3size;
        double sheetWidthMm = 0.0;
        double sheetHeightMm = 0.0;
        double scaleNumerator = 1.0;
        double scaleDenominator = 1.0;
        double frontXmm = 0.0;
        double frontYmm = 0.0;
        bool layoutFound = false;

        foreach (var sheet in sheets)
        {
            foreach (double candidateScaleDenominator in scaleDenominators)
            {
                double scale = 1.0 / candidateScaleDenominator;
                double frontWidthMm = part.BarLengthMm * scale;
                double frontHeightMm = part.BarHeightMm * scale;
                double totalWidthMm = leftMarginMm + frontWidthMm + sideMarginMm;
                double totalHeightMm = bottomTitleBlockClearanceMm + frontHeightMm + topMarginMm;

                if (totalWidthMm > sheet.WidthMm || totalHeightMm > sheet.HeightMm)
                    continue;

                paperCode = sheet.PaperCode;
                paperSize = sheet.PaperSize;
                sheetWidthMm = sheet.WidthMm;
                sheetHeightMm = sheet.HeightMm;
                scaleDenominator = candidateScaleDenominator;
                frontXmm = leftMarginMm + (frontWidthMm / 2.0);
                frontYmm = sheet.HeightMm - topMarginMm - (frontHeightMm / 2.0);
                layoutFound = true;
                break;
            }

            if (layoutFound)
                break;
        }

        if (!layoutFound)
            throw new InvalidOperationException("Could not find a drawing sheet size and scale that fits the Torsion Bar front view.");

        string defaultDrawingTemplate = string.Empty;
        try
        {
            defaultDrawingTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing) ?? string.Empty;
        }
        catch
        {
            defaultDrawingTemplate = string.Empty;
        }

        string drawingTemplatePath = string.Empty;
        if (!string.IsNullOrWhiteSpace(part.DrawingTemplatePathOverride))
        {
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
            string documentTemplateFolders = string.Empty;
            try
            {
                documentTemplateFolders = swApp.GetUserPreferenceStringListValue((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates) ?? string.Empty;
            }
            catch
            {
                documentTemplateFolders = string.Empty;
            }

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
            sheetFormatPath = Path.GetFullPath(part.DrawingSheetFormatPathOverride);
            if (!File.Exists(sheetFormatPath))
                throw new FileNotFoundException("The sheet-format override file was not found.", sheetFormatPath);
        }
        else
        {
            List<string> sheetFormatFolders = new();
            char[] separators = { '|', ';', '\r', '\n' };

            if (part.DrawingPreferSolidWorksTemplateLocations)
            {
                string configuredSheetFormatFolders = string.Empty;
                string configuredNewSheetFormatFolders = string.Empty;

                try
                {
                    configuredSheetFormatFolders = swApp.GetUserPreferenceStringListValue((int)swUserPreferenceStringValue_e.swFileLocationsSheetFormat) ?? string.Empty;
                }
                catch
                {
                    configuredSheetFormatFolders = string.Empty;
                }

                try
                {
                    configuredNewSheetFormatFolders = swApp.GetUserPreferenceStringListValue((int)swUserPreferenceStringValue_e.swFileLocationsNewSheetFormat) ?? string.Empty;
                }
                catch
                {
                    configuredNewSheetFormatFolders = string.Empty;
                }

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

        try
        {
            string sheetName = string.IsNullOrWhiteSpace(part.DrawingSheetName) ? "Sheet1" : part.DrawingSheetName;

            bool sheetConfigured = drawingDoc.SetupSheet6(
                sheetName,
                (int)paperSize,
                (int)swDwgTemplates_e.swDwgTemplateCustom,
                scaleNumerator,
                scaleDenominator,
                part.DrawingUseFirstAngleProjection,
                sheetFormatPath,
                Mm(sheetWidthMm),
                Mm(sheetHeightMm),
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

            drawingDoc.ActivateSheet(sheetName);
            drawingDoc.EditSheet();

            View frontView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Front", Mm(frontXmm), Mm(frontYmm), 0);
            if (frontView == null)
                throw new InvalidOperationException("Could not create the front torsion-bar view.");

            if (!string.IsNullOrWhiteSpace(part.DrawingReferencedConfiguration))
                frontView.ReferencedConfiguration = part.DrawingReferencedConfiguration;

            frontView.Position = new double[] { Mm(frontXmm), Mm(frontYmm), 0.0 };
            frontView.PositionLocked = false;

            drawingModel.EditRebuild3();
            drawingDoc.EditSheet();
            drawingDoc.ActivateSheet(sheetName);
            drawingModel.ClearSelection2(true);
            drawingModel.ViewZoomtofit2();

            string savedPath;
            if (part.DrawingSaveToPdm)
            {
                savedPath = pdm.SaveAsPdm(drawingModel, outFolder);
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
        catch
        {
            swApp.CloseDoc(drawingModel.GetTitle());
            throw;
        }
    }
}
