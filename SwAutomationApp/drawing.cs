using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation;

internal static class DrawingMethods
{
    public static string CreateTorsionBarDrawing(TorsionBarPart part, SldWorks swApp)
    {
        if (part == null)
            throw new InvalidOperationException("CreateTorsionBarDrawing requires a TorsionBarPart instance.");
        if (swApp == null)
            throw new ArgumentNullException(nameof(swApp));

        const double mmToMeters = AutomationSupport.MmToMeters;
        double Mm(double mm) => mm * mmToMeters;
        double MetersToLayoutMm(double meters) => meters / mmToMeters;

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
        string firstConfigurationName = AutomationSupport.RequireText(part.P0001ConfigName, nameof(part.P0001ConfigName), nameof(TorsionBarPart));
        string secondConfigurationName = AutomationSupport.RequireText(part.P0002ConfigName, nameof(part.P0002ConfigName), nameof(TorsionBarPart));
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
        const double leftMarginMm = 12.0;
        const double sideMarginMm = 12.0;
        const double topMarginMm = 14.0;
        const double sideViewGapMm = 18.0;
        const double configurationLabelBandHeightMm = 14.0;
        const double topDimensionBandMm = 16.0;
        const double lowerCalloutBandMm = 22.0;
        const double configurationRowGapMm = 24.0;
        const double configurationLabelTextHeightMm = 7.0;

        double bottomTitleBlockClearanceMm = Math.Max(0.0, MetersToLayoutMm(part.DrawingBottomTitleBlockClearanceMm));
        string paperCode = string.Empty;
        swDwgPaperSizes_e paperSize = swDwgPaperSizes_e.swDwgPaperA3size;
        double sheetHeightMm = 0.0;
        double scaleNumerator = 1.0;
        double scaleDenominator = 1.0;
        double firstFrontXmm = 0.0;
        double firstFrontYmm = 0.0;
        double firstSideXmm = 0.0;
        double firstSideYmm = 0.0;
        double secondFrontXmm = 0.0;
        double secondFrontYmm = 0.0;
        double secondSideXmm = 0.0;
        double secondSideYmm = 0.0;
        bool layoutFound = false;

        foreach (var sheet in sheets)
        {
            foreach (double candidateScaleDenominator in scaleDenominators)
            {
                double scale = 1.0 / candidateScaleDenominator;
                double frontWidthMm = MetersToLayoutMm(part.BarLengthMm) * scale;
                double frontHeightMm = MetersToLayoutMm(part.BarHeightMm) * scale;
                double sideWidthMm = MetersToLayoutMm(part.BarThicknessMm) * scale;
                double sideHeightMm = MetersToLayoutMm(part.BarHeightMm) * scale;
                double rowWidthMm = leftMarginMm + frontWidthMm + sideViewGapMm + sideWidthMm + sideMarginMm;
                double rowHeightMm = configurationLabelBandHeightMm + topDimensionBandMm + frontHeightMm + lowerCalloutBandMm;
                double totalHeightMm = topMarginMm + rowHeightMm + configurationRowGapMm + rowHeightMm + bottomTitleBlockClearanceMm;

                if (rowWidthMm > sheet.WidthMm || totalHeightMm > sheet.HeightMm)
                    continue;

                paperCode = sheet.PaperCode;
                paperSize = sheet.PaperSize;
                sheetHeightMm = sheet.HeightMm;
                scaleDenominator = candidateScaleDenominator;
                double firstRowTopYmm = sheet.HeightMm - topMarginMm;
                double secondRowTopYmm = firstRowTopYmm - rowHeightMm - configurationRowGapMm;

                firstFrontYmm = firstRowTopYmm - configurationLabelBandHeightMm - topDimensionBandMm - (frontHeightMm / 2.0);
                secondFrontYmm = secondRowTopYmm - configurationLabelBandHeightMm - topDimensionBandMm - (frontHeightMm / 2.0);
                firstSideYmm = firstFrontYmm;
                secondSideYmm = secondFrontYmm;

                if (part.DrawingUseFirstAngleProjection)
                {
                    firstSideXmm = leftMarginMm + (sideWidthMm / 2.0);
                    firstFrontXmm = firstSideXmm + (sideWidthMm / 2.0) + sideViewGapMm + (frontWidthMm / 2.0);
                    secondSideXmm = firstSideXmm;
                    secondFrontXmm = firstFrontXmm;
                }
                else
                {
                    firstFrontXmm = leftMarginMm + (frontWidthMm / 2.0);
                    firstSideXmm = firstFrontXmm + (frontWidthMm / 2.0) + sideViewGapMm + (sideWidthMm / 2.0);
                    secondFrontXmm = firstFrontXmm;
                    secondSideXmm = firstSideXmm;
                }

                layoutFound = true;
                break;
            }

            if (layoutFound)
                break;
        }

        if (!layoutFound)
            throw new InvalidOperationException("Could not find a drawing sheet size and scale that fits the torsion-bar configurations.");

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

        MathUtility mathUtility = swApp.GetMathUtility() as MathUtility;
        if (mathUtility == null)
        {
            swApp.CloseDoc(drawingModel.GetTitle());
            throw new InvalidOperationException("Could not access the SolidWorks math utility.");
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

            IEnumerable<object> EnumerateDispatchObjects(object raw)
            {
                if (raw == null)
                    return Array.Empty<object>();
                if (raw is object[] objectArray)
                    return objectArray;
                if (raw is Array anyArray)
                    return anyArray.Cast<object>();
                return new[] { raw };
            }

            double[] ToViewPointMm(View view, double x, double y, double z)
            {
                MathPoint modelPoint = mathUtility.CreatePoint(new double[] { x, y, z }) as MathPoint;
                if (modelPoint == null)
                    throw new InvalidOperationException("Could not create a math point for drawing geometry.");

                MathPoint viewPoint = modelPoint.MultiplyTransform(view.ModelToViewTransform) as MathPoint;
                double[] pointArray = viewPoint?.ArrayData as double[];
                if (pointArray == null || pointArray.Length < 2)
                    throw new InvalidOperationException($"Could not transform geometry into drawing coordinates for view '{view.Name}'.");

                return new[] { pointArray[0] / mmToMeters, pointArray[1] / mmToMeters };
            }

            List<(Edge Edge, double MidXMm, double MidYMm, bool IsHorizontal, bool IsVertical)> GetVisibleLineEdges(View view)
            {
                List<(Edge Edge, double MidXMm, double MidYMm, bool IsHorizontal, bool IsVertical)> lineEdges = new();
                foreach (object entityObject in EnumerateDispatchObjects(view.GetVisibleEntities2(null, (int)swViewEntityType_e.swViewEntityType_Edge)))
                {
                    Edge edge = entityObject as Edge;
                    Curve curve = edge?.GetCurve() as Curve;
                    if (edge == null || curve == null || !curve.IsLine())
                        continue;

                    Vertex startVertex = edge.GetStartVertex() as Vertex;
                    Vertex endVertex = edge.GetEndVertex() as Vertex;
                    double[] startPoint = startVertex?.GetPoint() as double[];
                    double[] endPoint = endVertex?.GetPoint() as double[];
                    if (startPoint == null || endPoint == null)
                        continue;

                    double[] viewStart = ToViewPointMm(view, startPoint[0], startPoint[1], startPoint[2]);
                    double[] viewEnd = ToViewPointMm(view, endPoint[0], endPoint[1], endPoint[2]);
                    bool isHorizontal = Math.Abs(viewStart[1] - viewEnd[1]) < 0.2;
                    bool isVertical = Math.Abs(viewStart[0] - viewEnd[0]) < 0.2;
                    lineEdges.Add((edge, (viewStart[0] + viewEnd[0]) / 2.0, (viewStart[1] + viewEnd[1]) / 2.0, isHorizontal, isVertical));
                }

                return lineEdges;
            }

            void SelectViewEntities(View view, params object[] entities)
            {
                drawingDoc.ActivateSheet(sheetName);
                drawingModel.ClearSelection2(true);

                for (int index = 0; index < entities.Length; index++)
                {
                    if (!view.SelectEntity(entities[index], index > 0))
                        throw new InvalidOperationException($"Could not select a drawing entity in view '{view.Name}'.");
                }
            }

            void FormatAnnotation(Annotation annotation, double textHeightMm, bool bold)
            {
                if (annotation == null)
                    return;

                TextFormat textFormat = annotation.GetTextFormat(0) as TextFormat;
                if (textFormat == null)
                    return;

                textFormat.CharHeight = Mm(textHeightMm);
                textFormat.Bold = bold;
                annotation.SetTextFormat(0, false, textFormat);
            }

            void AddPlainNote(string text, double xMm, double yMm, double textHeightMm, bool bold, int justification = (int)swTextJustification_e.swTextJustificationLeft)
            {
                drawingDoc.ActivateSheet(sheetName);
                drawingModel.ClearSelection2(true);

                Note note = drawingModel.InsertNote(text) as Note;
                if (note == null)
                    throw new InvalidOperationException($"Could not create drawing note '{text}'.");

                note.SetTextJustification(justification);
                Annotation annotation = note.GetAnnotation() as Annotation;
                if (annotation != null)
                {
                    annotation.SetPosition2(Mm(xMm), Mm(yMm), 0.0);
                    FormatAnnotation(annotation, textHeightMm, bold);
                }

                drawingModel.ClearSelection2(true);
            }

            void AddHorizontalDimension(View view, object firstEntity, object secondEntity, double xMm, double yMm, string description)
            {
                SelectViewEntities(view, firstEntity, secondEntity);
                DisplayDimension displayDimension = drawingModel.AddHorizontalDimension2(Mm(xMm), Mm(yMm), 0) as DisplayDimension;
                if (displayDimension == null)
                    throw new InvalidOperationException($"Could not create the horizontal dimension '{description}'.");
                drawingModel.ClearSelection2(true);
            }

            void AddVerticalDimension(View view, object firstEntity, object secondEntity, double xMm, double yMm, string description)
            {
                SelectViewEntities(view, firstEntity, secondEntity);
                DisplayDimension displayDimension = drawingModel.AddVerticalDimension2(Mm(xMm), Mm(yMm), 0) as DisplayDimension;
                if (displayDimension == null)
                    throw new InvalidOperationException($"Could not create the vertical dimension '{description}'.");
                drawingModel.ClearSelection2(true);
            }

            void DeleteCenterMarks(params View[] views)
            {
                List<CenterMark> centerMarks = new();
                foreach (View view in views)
                {
                    for (CenterMark centerMark = view.GetFirstCenterMark() as CenterMark; centerMark != null; centerMark = centerMark.GetNext() as CenterMark)
                        centerMarks.Add(centerMark);
                }

                if (centerMarks.Count == 0)
                    return;

                drawingModel.ClearSelection2(true);
                foreach (CenterMark centerMark in centerMarks)
                    centerMark.Select(true, null);
                drawingModel.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                drawingModel.ClearSelection2(true);
            }

            (double LeftMm, double BottomMm, double RightMm, double TopMm) GetViewOutlineMm(View view)
            {
                double[] outline = view.GetOutline() as double[];
                if (outline == null || outline.Length < 4)
                    throw new InvalidOperationException($"Could not read the outline of drawing view '{view.Name}'.");

                return (
                    outline[0] / mmToMeters,
                    outline[1] / mmToMeters,
                    outline[2] / mmToMeters,
                    outline[3] / mmToMeters);
            }

            void MoveViewCenter(View view, double centerXMm, double centerYMm)
            {
                view.Position = new double[] { Mm(centerXMm), Mm(centerYMm), 0.0 };
                view.PositionLocked = false;
            }

            void RepositionConfiguration(View frontView, View sideView, double targetFrontLeftMm, double targetFrontTopMm)
            {
                var frontOutline = GetViewOutlineMm(frontView);
                double frontWidthMm = frontOutline.RightMm - frontOutline.LeftMm;
                double frontHeightMm = frontOutline.TopMm - frontOutline.BottomMm;

                MoveViewCenter(frontView, targetFrontLeftMm + (frontWidthMm / 2.0), targetFrontTopMm - (frontHeightMm / 2.0));
                drawingModel.EditRebuild3();

                frontOutline = GetViewOutlineMm(frontView);
                var sideOutline = GetViewOutlineMm(sideView);
                double sideWidthMm = sideOutline.RightMm - sideOutline.LeftMm;
                double sideHeightMm = sideOutline.TopMm - sideOutline.BottomMm;
                // Tune this to change the horizontal gap between the front and side view.
                double targetSideLeftMm = frontOutline.RightMm + sideViewGapMm + 10.0;
                double targetSideCenterYmm = frontOutline.BottomMm + (frontHeightMm / 2.0);
                // Tune this to move the side view slightly up or down relative to the front view.
                double sideVerticalNudgeMm = (frontHeightMm - sideHeightMm) / 2.0;

                MoveViewCenter(sideView, targetSideLeftMm + (sideWidthMm / 2.0), targetSideCenterYmm + sideVerticalNudgeMm);
                drawingModel.EditRebuild3();
            }

            drawingDoc.ActivateSheet(sheetName);
            drawingDoc.EditSheet();
            drawingModel.EditRebuild3();

            View firstFrontView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Front", Mm(firstFrontXmm), Mm(firstFrontYmm), 0);
            if (firstFrontView == null)
                throw new InvalidOperationException("Could not create the P0001 front torsion-bar view.");

            firstFrontView.ReferencedConfiguration = firstConfigurationName;
            firstFrontView.Position = new double[] { Mm(firstFrontXmm), Mm(firstFrontYmm), 0.0 };
            firstFrontView.PositionLocked = false;
            drawingModel.EditRebuild3();

            drawingModel.ClearSelection2(true);
            if (!drawingDoc.ActivateView(firstFrontView.Name))
                throw new InvalidOperationException("Could not activate the P0001 front torsion-bar view before creating the side view.");
            if (!drawingModel.Extension.SelectByID2(firstFrontView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0))
                throw new InvalidOperationException("Could not select the P0001 front torsion-bar view before creating the side view.");

            View firstSideView = drawingDoc.CreateUnfoldedViewAt3(Mm(firstSideXmm), Mm(firstSideYmm), 0, false);
            if (firstSideView == null)
                throw new InvalidOperationException("Could not create the projected P0001 side torsion-bar view.");

            firstSideView.PositionLocked = false;

            drawingDoc.ActivateSheet(sheetName);
            drawingModel.ClearSelection2(true);

            View secondFrontView = drawingDoc.CreateDrawViewFromModelView3(partPath, "*Front", Mm(secondFrontXmm), Mm(secondFrontYmm), 0);
            if (secondFrontView == null)
                throw new InvalidOperationException("Could not create the P0002 front torsion-bar view.");

            secondFrontView.ReferencedConfiguration = secondConfigurationName;
            secondFrontView.Position = new double[] { Mm(secondFrontXmm), Mm(secondFrontYmm), 0.0 };
            secondFrontView.PositionLocked = false;
            drawingModel.EditRebuild3();

            drawingModel.ClearSelection2(true);
            if (!drawingDoc.ActivateView(secondFrontView.Name))
                throw new InvalidOperationException("Could not activate the P0002 front torsion-bar view before creating the side view.");
            if (!drawingModel.Extension.SelectByID2(secondFrontView.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0))
                throw new InvalidOperationException("Could not select the P0002 front torsion-bar view before creating the side view.");

            View secondSideView = drawingDoc.CreateUnfoldedViewAt3(Mm(secondSideXmm), Mm(secondSideYmm), 0, false);
            if (secondSideView == null)
                throw new InvalidOperationException("Could not create the projected P0002 side torsion-bar view.");

            secondSideView.PositionLocked = false;
            drawingModel.EditRebuild3();

            // Tune these two numbers to move the full P0001 view block: left position, then top position.
            RepositionConfiguration(firstFrontView, firstSideView, 24.0, sheetHeightMm - 64.0);
            // Tune these two numbers to move the full P0002 view block: left position, then top position.
            RepositionConfiguration(secondFrontView, secondSideView, 24.0, sheetHeightMm - 154.0);
            DeleteCenterMarks(firstFrontView, firstSideView);

            var firstFrontOutline = GetViewOutlineMm(firstFrontView);
            var firstSideOutline = GetViewOutlineMm(firstSideView);
            var secondFrontOutline = GetViewOutlineMm(secondFrontView);
            var secondSideOutline = GetViewOutlineMm(secondSideView);

            // Tune `+ 16.0` to move the P0001 label up or down.
            AddPlainNote(firstConfigurationName, firstFrontOutline.LeftMm, firstFrontOutline.TopMm + 16.0, configurationLabelTextHeightMm, true);
            // Tune `+ 16.0` to move the P0002 label up or down.
            AddPlainNote(secondConfigurationName, secondFrontOutline.LeftMm, secondFrontOutline.TopMm + 16.0, configurationLabelTextHeightMm, true);

            List<(Edge Edge, double MidXMm, double MidYMm, bool IsHorizontal, bool IsVertical)> firstFrontLines = GetVisibleLineEdges(firstFrontView);
            List<(Edge Edge, double MidXMm, double MidYMm, bool IsHorizontal, bool IsVertical)> firstSideLines = GetVisibleLineEdges(firstSideView);
            List<(Edge Edge, double MidXMm, double MidYMm, bool IsHorizontal, bool IsVertical)> secondFrontLines = GetVisibleLineEdges(secondFrontView);
            List<(Edge Edge, double MidXMm, double MidYMm, bool IsHorizontal, bool IsVertical)> secondSideLines = GetVisibleLineEdges(secondSideView);

            Edge firstFrontLeftEdge = firstFrontLines.Where(line => line.IsVertical).OrderBy(line => line.MidXMm).First().Edge;
            Edge firstFrontRightEdge = firstFrontLines.Where(line => line.IsVertical).OrderBy(line => line.MidXMm).Last().Edge;
            Edge firstSideLeftEdge = firstSideLines.Where(line => line.IsVertical).OrderBy(line => line.MidXMm).First().Edge;
            Edge firstSideRightEdge = firstSideLines.Where(line => line.IsVertical).OrderBy(line => line.MidXMm).Last().Edge;
            Edge firstSideBottomEdge = firstSideLines.Where(line => line.IsHorizontal).OrderBy(line => line.MidYMm).First().Edge;
            Edge firstSideTopEdge = firstSideLines.Where(line => line.IsHorizontal).OrderBy(line => line.MidYMm).Last().Edge;

            Edge secondFrontLeftEdge = secondFrontLines.Where(line => line.IsVertical).OrderBy(line => line.MidXMm).First().Edge;
            Edge secondFrontRightEdge = secondFrontLines.Where(line => line.IsVertical).OrderBy(line => line.MidXMm).Last().Edge;
            Edge secondFrontBottomEdge = secondFrontLines.Where(line => line.IsHorizontal).OrderBy(line => line.MidYMm).First().Edge;
            Edge secondSideLeftEdge = secondSideLines.Where(line => line.IsVertical).OrderBy(line => line.MidXMm).First().Edge;
            Edge secondSideRightEdge = secondSideLines.Where(line => line.IsVertical).OrderBy(line => line.MidXMm).Last().Edge;
            Edge secondSideBottomEdge = secondSideLines.Where(line => line.IsHorizontal).OrderBy(line => line.MidYMm).First().Edge;
            Edge secondSideTopEdge = secondSideLines.Where(line => line.IsHorizontal).OrderBy(line => line.MidYMm).Last().Edge;

            // Tune `+ 8.0` to move the P0001 length dimension up or down.
            AddHorizontalDimension(firstFrontView, firstFrontLeftEdge, firstFrontRightEdge,
                (firstFrontOutline.LeftMm + firstFrontOutline.RightMm) / 2.0,
                firstFrontOutline.TopMm + 8.0,
                "P0001 overall length");

            // Tune `+ 8.0` to move the P0001 thickness dimension up or down.
            AddHorizontalDimension(firstSideView, firstSideLeftEdge, firstSideRightEdge,
                (firstSideOutline.LeftMm + firstSideOutline.RightMm) / 2.0,
                firstSideOutline.TopMm + 8.0,
                "P0001 side thickness");

            // Tune `+ 7.0` to move the P0001 height dimension farther right or closer to the view.
            AddVerticalDimension(firstSideView, firstSideBottomEdge, firstSideTopEdge,
                firstSideOutline.RightMm + 15.0,
                (firstSideOutline.BottomMm + firstSideOutline.TopMm) / 2.0,
                "P0001 side height");

            // Tune `+ 12.0` to move the P0002 length dimension up or down.
            AddHorizontalDimension(secondFrontView, secondFrontLeftEdge, secondFrontRightEdge,
                (secondFrontOutline.LeftMm + secondFrontOutline.RightMm) / 2.0,
                secondFrontOutline.TopMm + 12.0,
                "P0002 overall length");

            // Tune `+ 8.0` to move the P0002 thickness dimension up or down.
            AddHorizontalDimension(secondSideView, secondSideLeftEdge, secondSideRightEdge,
                (secondSideOutline.LeftMm + secondSideOutline.RightMm) / 2.0,
                secondSideOutline.TopMm + 8.0,
                "P0002 side thickness");

            // Tune `+ 7.0` to move the P0002 height dimension farther right or closer to the view.
            AddVerticalDimension(secondSideView, secondSideBottomEdge, secondSideTopEdge,
                secondSideOutline.RightMm + 15.0,
                (secondSideOutline.BottomMm + secondSideOutline.TopMm) / 2.0,
                "P0002 side height");

            drawingModel.EditRebuild3();
            drawingDoc.ActivateSheet(sheetName);
            drawingModel.ClearSelection2(true);
            drawingModel.ViewZoomtofit2();

            string savedPath = Path.Combine(outFolder, drawingFileName);
            drawingModel.SaveAs3(savedPath, 0, 1);
            Console.WriteLine($"Drawing saved locally: {savedPath}");

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
