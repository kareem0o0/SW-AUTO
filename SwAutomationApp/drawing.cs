using System;
using System.Collections.Generic;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

public sealed class TorsionBarDrawing
{
    private const double MmToMeters = AutomationSupport.MmToMeters;
    private const string DefaultBirrSheetFormatFolder = @"C:\Users\kareem.salah\PDM\Birr Machines PDM\40_Templates\Solidworks\Blattformate\Birr Machines";

    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    private readonly struct SheetOption
    {
        public SheetOption(string paperCode, swDwgPaperSizes_e paperSize, double widthMm, double heightMm)
        {
            PaperCode = paperCode;
            PaperSize = paperSize;
            WidthMm = widthMm;
            HeightMm = heightMm;
        }

        public string PaperCode { get; }
        public swDwgPaperSizes_e PaperSize { get; }
        public double WidthMm { get; }
        public double HeightMm { get; }
    }

    private readonly struct LayoutPlan
    {
        public LayoutPlan(
            SheetOption sheet,
            double scaleNumerator,
            double scaleDenominator,
            double frontXmm,
            double frontYmm,
            double topXmm,
            double topYmm,
            double rightXmm,
            double rightYmm)
        {
            Sheet = sheet;
            ScaleNumerator = scaleNumerator;
            ScaleDenominator = scaleDenominator;
            FrontXmm = frontXmm;
            FrontYmm = frontYmm;
            TopXmm = topXmm;
            TopYmm = topYmm;
            RightXmm = rightXmm;
            RightYmm = rightYmm;
        }

        public SheetOption Sheet { get; }
        public double ScaleNumerator { get; }
        public double ScaleDenominator { get; }
        public double FrontXmm { get; }
        public double FrontYmm { get; }
        public double TopXmm { get; }
        public double TopYmm { get; }
        public double RightXmm { get; }
        public double RightYmm { get; }
    }

    public TorsionBarDrawing(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        _pdm = pdm ?? throw new ArgumentNullException(nameof(pdm));
    }

    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string LocalFileName { get; set; } = "TorsionBar.SLDDRW";
    public string SheetName { get; set; } = "Sheet1";
    public string LanguageCode { get; set; } = "EN";
    public string TemplateFolderPath { get; set; } = DefaultBirrSheetFormatFolder;
    public bool PreferSolidWorksTemplateLocations { get; set; } = true;
    public string SheetFormatPathOverride { get; set; } = string.Empty;
    public string DrawingTemplatePathOverride { get; set; } = string.Empty;
    public bool UseFirstAngleProjection { get; set; }
    public double BottomTitleBlockClearanceMm { get; set; } = 85.0;
    public string ReferencedConfiguration { get; set; } = "P0002";
    public TorsionBarPart Part { get; set; } = null;

    private string GetRequiredOutputFolder() => AutomationSupport.RequireText(OutputFolder, nameof(OutputFolder), nameof(TorsionBarDrawing));
    private string GetRequiredLocalFileName() => AutomationSupport.RequireText(LocalFileName, nameof(LocalFileName), nameof(TorsionBarDrawing));
    private TorsionBarPart GetRequiredPart()
    {
        if (Part == null)
            throw new InvalidOperationException("TorsionBarDrawing requires a TorsionBarPart instance.");

        return Part;
    }

    public string Create()
    {
        string outFolder = GetRequiredOutputFolder();
        Directory.CreateDirectory(outFolder);
        TorsionBarPart part = GetRequiredPart();

        string partFileName = part.Create();
        if (string.IsNullOrWhiteSpace(partFileName))
            throw new InvalidOperationException("TorsionBarDrawing could not create the source TorsionBarPart.");

        string partPath = Path.GetFullPath(Path.Combine(part.OutputFolder, partFileName));
        if (!File.Exists(partPath))
            throw new FileNotFoundException("The source Torsion Bar part file was not found after creation.", partPath);

        LayoutPlan layout = BuildLayoutPlan(part);
        string drawingTemplatePath = ResolveDrawingTemplatePath();
        string sheetFormatPath = ResolveSheetFormatPath(layout.Sheet.PaperCode);

        ModelDoc2 drawingModel = (ModelDoc2)_swApp.NewDocument(drawingTemplatePath, 0, 0, 0);
        if (drawingModel == null)
            throw new InvalidOperationException("Failed to create a new drawing document.");

        DrawingDoc drawingDoc = drawingModel as DrawingDoc;
        if (drawingDoc == null)
        {
            _swApp.CloseDoc(drawingModel.GetTitle());
            throw new InvalidOperationException("Could not access the SolidWorks drawing document.");
        }

        try
        {
            string sheetName = string.IsNullOrWhiteSpace(SheetName) ? "Sheet1" : SheetName;

            bool sheetConfigured = drawingDoc.SetupSheet6(
                sheetName,
                (int)layout.Sheet.PaperSize,
                (int)swDwgTemplates_e.swDwgTemplateCustom,
                layout.ScaleNumerator,
                layout.ScaleDenominator,
                UseFirstAngleProjection,
                sheetFormatPath,
                Mm(layout.Sheet.WidthMm),
                Mm(layout.Sheet.HeightMm),
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

            View frontView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Front", Mm(layout.FrontXmm), Mm(layout.FrontYmm), 0);
            View topView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Top", Mm(layout.TopXmm), Mm(layout.TopYmm), 0);
            View rightView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Right", Mm(layout.RightXmm), Mm(layout.RightYmm), 0);

            if (frontView == null || topView == null || rightView == null)
                throw new InvalidOperationException("Could not create the required drawing views.");

            if (!string.IsNullOrWhiteSpace(ReferencedConfiguration))
            {
                frontView.ReferencedConfiguration = ReferencedConfiguration;
                topView.ReferencedConfiguration = ReferencedConfiguration;
                rightView.ReferencedConfiguration = ReferencedConfiguration;
            }

            // Front view carries the hole pattern and overall length. Right view captures the thickness.
            TryAutoDimensionView(drawingDoc, frontView.Name);
            TryAutoDimensionView(drawingDoc, rightView.Name);

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
            if (SaveToPdm)
            {
                savedPath = _pdm.SaveAsPdm(drawingModel, outFolder);
                Console.WriteLine($"Drawing saved to PDM: {savedPath}");
            }
            else
            {
                savedPath = Path.Combine(outFolder, GetRequiredLocalFileName());
                drawingModel.SaveAs3(savedPath, 0, 1);
                Console.WriteLine($"Drawing saved locally: {savedPath}");
            }

            if (CloseAfterCreate)
            {
                _swApp.CloseDoc(drawingModel.GetTitle());
                Console.WriteLine("Drawing closed after creating.");
            }

            return Path.GetFileName(savedPath);
        }
        catch
        {
            _swApp.CloseDoc(drawingModel.GetTitle());
            throw;
        }

        double Mm(double mm) => mm * MmToMeters;
    }

    private string ResolveDrawingTemplatePath()
    {
        if (!string.IsNullOrWhiteSpace(DrawingTemplatePathOverride))
        {
            string overridePath = Path.GetFullPath(DrawingTemplatePathOverride);
            if (File.Exists(overridePath))
                return overridePath;

            throw new FileNotFoundException("The drawing template override file was not found.", overridePath);
        }

        string defaultTemplate = TryGetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
        if (!string.IsNullOrWhiteSpace(defaultTemplate) && File.Exists(defaultTemplate))
            return defaultTemplate;

        foreach (string folder in GetConfiguredSearchFolders((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates))
        {
            string[] templates = Directory.GetFiles(folder, "*.drwdot", SearchOption.TopDirectoryOnly);
            if (templates.Length > 0)
                return templates[0];
        }

        throw new InvalidOperationException("Could not resolve a SolidWorks drawing document template. Set DrawingTemplatePathOverride or configure a default drawing template in SolidWorks.");
    }

    private string ResolveSheetFormatPath(string paperCode)
    {
        if (!string.IsNullOrWhiteSpace(SheetFormatPathOverride))
        {
            string overridePath = Path.GetFullPath(SheetFormatPathOverride);
            if (File.Exists(overridePath))
                return overridePath;

            throw new FileNotFoundException("The sheet-format override file was not found.", overridePath);
        }

        string normalizedLanguageCode = NormalizeLanguageCode();
        string expectedFileName = $"{paperCode}_Birr_Machines_{normalizedLanguageCode}.slddrt";

        foreach (string folder in GetSheetFormatSearchFolders())
        {
            string directPath = Path.Combine(folder, expectedFileName);
            if (File.Exists(directPath))
                return directPath;

            string[] recursiveMatches = Directory.GetFiles(folder, expectedFileName, SearchOption.AllDirectories);
            if (recursiveMatches.Length > 0)
                return recursiveMatches[0];
        }

        throw new FileNotFoundException($"Could not find sheet format '{expectedFileName}'. Check LanguageCode or TemplateFolderPath.");
    }

    private IEnumerable<string> GetSheetFormatSearchFolders()
    {
        HashSet<string> folders = new(StringComparer.OrdinalIgnoreCase);

        if (PreferSolidWorksTemplateLocations)
        {
            foreach (string folder in GetConfiguredSearchFolders((int)swUserPreferenceStringValue_e.swFileLocationsSheetFormat))
                folders.Add(folder);

            foreach (string folder in GetConfiguredSearchFolders((int)swUserPreferenceStringValue_e.swFileLocationsNewSheetFormat))
                folders.Add(folder);

            string defaultDrawingTemplate = TryGetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
            if (!string.IsNullOrWhiteSpace(defaultDrawingTemplate) && File.Exists(defaultDrawingTemplate))
            {
                string templateFolder = Path.GetDirectoryName(defaultDrawingTemplate) ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(templateFolder) && Directory.Exists(templateFolder))
                    folders.Add(templateFolder);
            }
        }

        if (!string.IsNullOrWhiteSpace(TemplateFolderPath) && Directory.Exists(TemplateFolderPath))
            folders.Add(Path.GetFullPath(TemplateFolderPath));

        return folders;
    }

    private IEnumerable<string> GetConfiguredSearchFolders(int preferenceId)
    {
        string configuredFolderList = TryGetUserPreferenceStringListValue(preferenceId);
        if (string.IsNullOrWhiteSpace(configuredFolderList))
            yield break;

        char[] separators = { '|', ';', '\r', '\n' };
        foreach (string rawFolder in configuredFolderList.Split(separators, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (Directory.Exists(rawFolder))
                yield return rawFolder;
        }
    }

    private string TryGetUserPreferenceStringValue(int preferenceId)
    {
        try
        {
            return _swApp.GetUserPreferenceStringValue(preferenceId) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private string TryGetUserPreferenceStringListValue(int preferenceId)
    {
        try
        {
            return _swApp.GetUserPreferenceStringListValue(preferenceId) ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private string NormalizeLanguageCode()
    {
        string normalized = (LanguageCode ?? string.Empty).Trim().ToUpperInvariant();
        if (normalized == "EN" || normalized == "DE")
            return normalized;

        throw new InvalidOperationException("TorsionBarDrawing.LanguageCode must be either EN or DE.");
    }

    private void TryAutoDimensionView(DrawingDoc drawingDoc, string viewName)
    {
        if (string.IsNullOrWhiteSpace(viewName))
            return;

        if (!drawingDoc.ActivateView(viewName))
            return;

        int autoDimensionStatus = drawingDoc.AutoDimension(
            (int)swAutodimEntities_e.swAutodimEntitiesAll,
            (int)swAutodimScheme_e.swAutodimSchemeBaseline,
            (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementAbove,
            (int)swAutodimScheme_e.swAutodimSchemeBaseline,
            (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementRight);

        Console.WriteLine($"Auto-dimension status for {viewName}: {autoDimensionStatus}");
    }

    private LayoutPlan BuildLayoutPlan(TorsionBarPart part)
    {
        SheetOption[] sheets =
        {
            new("A4", swDwgPaperSizes_e.swDwgPaperA4size, 297.0, 210.0),
            new("A3", swDwgPaperSizes_e.swDwgPaperA3size, 420.0, 297.0),
            new("A2", swDwgPaperSizes_e.swDwgPaperA2size, 594.0, 420.0),
            new("A1", swDwgPaperSizes_e.swDwgPaperA1size, 841.0, 594.0)
        };

        double[] scaleDenominators = { 1.0, 2.0, 3.0, 4.0, 5.0, 10.0 };
        const double sideMarginMm = 20.0;
        const double topMarginMm = 20.0;
        const double viewGapMm = 18.0;
        const double rightViewReserveMm = 45.0;
        double bottomTitleBlockClearanceMm = Math.Max(0.0, BottomTitleBlockClearanceMm);

        foreach (SheetOption sheet in sheets)
        {
            foreach (double scaleDenominator in scaleDenominators)
            {
                double scale = 1.0 / scaleDenominator;
                double frontWidthMm = part.BarLengthMm * scale;
                double frontHeightMm = part.BarHeightMm * scale;
                double topHeightMm = part.BarThicknessMm * scale;
                double rightWidthMm = part.BarThicknessMm * scale;
                double rightHeightMm = part.BarHeightMm * scale;
                double totalWidthMm = sideMarginMm + frontWidthMm + viewGapMm + Math.Max(rightWidthMm, rightViewReserveMm) + sideMarginMm;
                double totalHeightMm = bottomTitleBlockClearanceMm + topHeightMm + viewGapMm + Math.Max(frontHeightMm, rightHeightMm) + topMarginMm;

                if (totalWidthMm > sheet.WidthMm || totalHeightMm > sheet.HeightMm)
                    continue;

                double frontXmm = sideMarginMm + (frontWidthMm / 2.0);
                double topXmm = frontXmm;
                double rightXmm = sideMarginMm + frontWidthMm + viewGapMm + (rightWidthMm / 2.0);
                double topYmm = sheet.HeightMm - topMarginMm - (topHeightMm / 2.0);
                double frontYmm = topYmm - (topHeightMm / 2.0) - viewGapMm - (frontHeightMm / 2.0);
                double rightYmm = frontYmm;

                return new LayoutPlan(
                    sheet,
                    1.0,
                    scaleDenominator,
                    frontXmm,
                    frontYmm,
                    topXmm,
                    topYmm,
                    rightXmm,
                    rightYmm);
            }
        }

        throw new InvalidOperationException("Could not find a drawing sheet size and scale that fits the Torsion Bar views.");
    }
}
