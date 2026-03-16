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

        // Use the drawing folder if one was provided. Otherwise save the drawing next to the part.
        string outFolder = string.IsNullOrWhiteSpace(part.DrawingOutputFolder)
            ? AutomationSupport.RequireText(part.OutputFolder, nameof(part.OutputFolder), nameof(TorsionBarPart))
            : Path.GetFullPath(part.DrawingOutputFolder);
        Directory.CreateDirectory(outFolder);

        string drawingFileName = string.IsNullOrWhiteSpace(part.DrawingLocalFileName)
            ? Path.ChangeExtension(part.LocalFileName, ".SLDDRW")
            : part.DrawingLocalFileName;

        // The part class still owns part creation. The drawing method just uses it.
        string partFileName = part.CreatePart();
        if (string.IsNullOrWhiteSpace(partFileName))
            throw new InvalidOperationException("CreateTorsionBarDrawing could not create the source TorsionBarPart.");

        string partPath = Path.GetFullPath(Path.Combine(part.OutputFolder, partFileName));
        if (!File.Exists(partPath))
            throw new FileNotFoundException("The source Torsion Bar part file was not found after creation.", partPath);

        string languageCode = (part.DrawingLanguageCode ?? string.Empty).Trim().ToUpperInvariant();
        if (languageCode != "EN" && languageCode != "DE")
            throw new InvalidOperationException("TorsionBarPart.DrawingLanguageCode must be either EN or DE.");

        // Pick the smallest paper size and scale that can fit the standard views above the title block.
        (string PaperCode, swDwgPaperSizes_e PaperSize, double WidthMm, double HeightMm)[] sheets =
        {
            ("A4", swDwgPaperSizes_e.swDwgPaperA4size, 297.0, 210.0),
            ("A3", swDwgPaperSizes_e.swDwgPaperA3size, 420.0, 297.0),
            ("A2", swDwgPaperSizes_e.swDwgPaperA2size, 594.0, 420.0),
            ("A1", swDwgPaperSizes_e.swDwgPaperA1size, 841.0, 594.0)
        };
        double[] scaleDenominators = { 1.0, 2.0, 3.0, 4.0, 5.0, 10.0 };
        const double sideMarginMm = 20.0;
        const double topMarginMm = 20.0;
        const double viewGapMm = 18.0;
        const double rightViewReserveMm = 45.0;
        double bottomTitleBlockClearanceMm = Math.Max(0.0, part.DrawingBottomTitleBlockClearanceMm);

        string paperCode = string.Empty;
        swDwgPaperSizes_e paperSize = swDwgPaperSizes_e.swDwgPaperA3size;
        double sheetWidthMm = 0.0;
        double sheetHeightMm = 0.0;
        double scaleNumerator = 1.0;
        double scaleDenominator = 1.0;
        double frontXmm = 0.0;
        double frontYmm = 0.0;
        double topXmm = 0.0;
        double topYmm = 0.0;
        double rightXmm = 0.0;
        double rightYmm = 0.0;
        bool layoutFound = false;

        foreach (var sheet in sheets)
        {
            foreach (double candidateScaleDenominator in scaleDenominators)
            {
                double scale = 1.0 / candidateScaleDenominator;
                double frontWidthMm = part.BarLengthMm * scale;
                double frontHeightMm = part.BarHeightMm * scale;
                double topHeightMm = part.BarThicknessMm * scale;
                double rightWidthMm = part.BarThicknessMm * scale;
                double rightHeightMm = part.BarHeightMm * scale;
                double totalWidthMm = sideMarginMm + frontWidthMm + viewGapMm + Math.Max(rightWidthMm, rightViewReserveMm) + sideMarginMm;
                double totalHeightMm = bottomTitleBlockClearanceMm + topHeightMm + viewGapMm + Math.Max(frontHeightMm, rightHeightMm) + topMarginMm;

                if (totalWidthMm > sheet.WidthMm || totalHeightMm > sheet.HeightMm)
                    continue;

                paperCode = sheet.PaperCode;
                paperSize = sheet.PaperSize;
                sheetWidthMm = sheet.WidthMm;
                sheetHeightMm = sheet.HeightMm;
                scaleDenominator = candidateScaleDenominator;
                topXmm = sideMarginMm + (frontWidthMm / 2.0);
                frontXmm = topXmm;
                rightXmm = sideMarginMm + frontWidthMm + viewGapMm + (rightWidthMm / 2.0);
                topYmm = sheet.HeightMm - topMarginMm - (topHeightMm / 2.0);
                frontYmm = topYmm - (topHeightMm / 2.0) - viewGapMm - (frontHeightMm / 2.0);
                rightYmm = frontYmm;
                layoutFound = true;
                break;
            }

            if (layoutFound)
                break;
        }

        if (!layoutFound)
            throw new InvalidOperationException("Could not find a drawing sheet size and scale that fits the Torsion Bar views.");

        // Try the SolidWorks document template first, then the configured template folders.
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

        // Build the sheet-format search list:
        // SolidWorks-configured folders first, then the Birr template folder on the part object.
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

            View frontView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Front", Mm(frontXmm), Mm(frontYmm), 0);
            View topView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Top", Mm(topXmm), Mm(topYmm), 0);
            View rightView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Right", Mm(rightXmm), Mm(rightYmm), 0);

            if (frontView == null || topView == null || rightView == null)
                throw new InvalidOperationException("Could not create the required drawing views.");

            if (!string.IsNullOrWhiteSpace(part.DrawingReferencedConfiguration))
            {
                frontView.ReferencedConfiguration = part.DrawingReferencedConfiguration;
                topView.ReferencedConfiguration = part.DrawingReferencedConfiguration;
                rightView.ReferencedConfiguration = part.DrawingReferencedConfiguration;
            }

            if (drawingDoc.ActivateView(frontView.Name))
            {
                drawingDoc.AutoDimension(
                    (int)swAutodimEntities_e.swAutodimEntitiesAll,
                    (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                    (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementAbove,
                    (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                    (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementRight);
            }

            if (drawingDoc.ActivateView(rightView.Name))
            {
                drawingDoc.AutoDimension(
                    (int)swAutodimEntities_e.swAutodimEntitiesAll,
                    (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                    (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementAbove,
                    (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                    (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementRight);
            }

            drawingDoc.ActivateView(frontView.Name);
            drawingDoc.InsertModelAnnotations3(
                (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing
                | (int)swInsertAnnotation_e.swInsertDimensionsNotMarkedForDrawing
                | (int)swInsertAnnotation_e.swInsertHoleWizardProfileDimensions
                | (int)swInsertAnnotation_e.swInsertHoleWizardLocationDimensions
                | (int)swInsertAnnotation_e.swInsertholeCallout,
                false,
                false,
                false,
                true);

            drawingModel.EditRebuild3();
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
