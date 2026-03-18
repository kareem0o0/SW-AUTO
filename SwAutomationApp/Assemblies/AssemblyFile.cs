using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

/// <summary>
/// Generic assembly document helper.
///
/// This class is responsible for:
/// 1. creating an assembly document
/// 2. saving it locally or to PDM
/// 3. inserting components
/// 4. creating common mate types
/// 5. creating linear and circular component patterns
///
/// It does not decide which parts should be inserted or how a full machine is built.
/// That orchestration stays in macros such as Run4().
/// </summary>
public class AssemblyFile
{
    // Shared application/service references passed in from the outside world.
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;

    // These fields track the currently open assembly document while insert/mate operations run.
    private ModelDoc2 _model = null;
    private AssemblyDoc _assembly = null;
    private string _folder = string.Empty;
    private int _insertIndex = 0;
    private Component2 _lastInsertedComponent = null;

    public AssemblyFile(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp;
        _pdm = pdm;
    }

    // These properties define how the assembly document should be saved and managed.
    public string OutputFolder { get; set; } = string.Empty;
    public bool CloseAfterCreate { get; set; }
    public bool SaveToPdm { get; set; }
    public string FileName { get; set; } = "Assembly.SLDASM";
    public BirrDataCardValues PdmDataCard { get; set; } = BirrDataCardValues.CreateDefault();

    private string GetRequiredOutputFolder()
    {
        return OutputFolder;
    }
    private string GetRequiredFileName()
    {
        return FileName;
    }

    /// <summary>
    /// Public entry point used by macros and external callers.
    /// </summary>
    public string Create()
    {
        return CreateDocument(CloseAfterCreate);
    }

    private string CreateDocument(bool closeAfterCreate)
    {
        // Read and validate the current object state first.
        string outFolder = GetRequiredOutputFolder();
        string fileName = GetRequiredFileName();
        bool saveToPdm = SaveToPdm;

        Directory.CreateDirectory(outFolder);

        // Ask SolidWorks for the user's default assembly template,
        // then create a new empty assembly document from it.
        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
        ModelDoc2 model = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
        string fullPath;

        // Save either to PDM or to a normal local file path.
        if (saveToPdm)
        {
            fullPath = _pdm.SaveAsPdm(model, outFolder, PdmDataCard);
        }
        else
        {
            fullPath = Path.Combine(outFolder, fileName);
            model.SaveAs3(fullPath, 0, 1);
        }

        Console.WriteLine($"Assembly saved to: {fullPath}");
        _insertIndex = 0;

        if (closeAfterCreate)
        {
            // Some flows only need the file to exist on disk.
            // In that case we close the assembly immediately and return the file name.
            _swApp.CloseDoc(model.GetTitle());
            _model = null;
            _assembly = null;
            _folder = string.Empty;
            _lastInsertedComponent = null;
            Console.WriteLine("Assembly closed after creating.");
            return Path.GetFileName(fullPath);
        }

        // Keep the assembly open when later insert/mate operations are expected.
        _model = model;
        _assembly = model as AssemblyDoc;
        _folder = Path.GetDirectoryName(fullPath) ?? string.Empty;

        int activateErrors = 0;
        _swApp.ActivateDoc3(_model.GetTitle(), false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref activateErrors);
        Console.WriteLine("Assembly remains open for further operations.");
        return Path.GetFileName(fullPath);
    }

    /// <summary>
    /// Closes the currently open assembly document and clears the cached state.
    /// </summary>
    public void Close()
    {
        if (_model == null)
            return;

        _swApp.CloseDoc(_model.GetTitle());
        _model = null;
        _assembly = null;
        _folder = string.Empty;
        _insertIndex = 0;
        _lastInsertedComponent = null;
    }

    /// <summary>
    /// Inserts a saved part or assembly file into the currently open assembly.
    ///
    /// The input may be a full path or a file name relative to the assembly folder.
    /// </summary>
    public Component2 Insert(string componentPath)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        if (string.IsNullOrWhiteSpace(componentPath))
            throw new ArgumentException("Component path is required.", nameof(componentPath));

        // Resolve a usable full path first.
        string path = Path.GetFullPath(
            Path.IsPathRooted(componentPath)
                ? componentPath
                : Path.Combine(_folder, componentPath));

        // Decide whether we are inserting a part file or another assembly file.
        int docType = Path.GetExtension(path)
            .Equals(".sldasm", StringComparison.OrdinalIgnoreCase)
                ? (int)swDocumentTypes_e.swDocASSEMBLY
                : (int)swDocumentTypes_e.swDocPART;

        // Open the source document silently so SolidWorks can insert it reliably.
        int openErrors = 0;
        int openWarnings = 0;
        int activateErrors = 0;

        ModelDoc2 openedModel = _swApp.OpenDoc6(
            path,
            docType,
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
            string.Empty,
            ref openErrors,
            ref openWarnings);

        // Reactivate the target assembly before inserting the component.
        _swApp.ActivateDoc3(
            _model.GetTitle(),
            false,
            (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
            ref activateErrors);

        // AddComponent5 inserts the document into the assembly space.
        // A small X offset is used while inserting multiple components so they do not land
        // exactly on top of each other before mates are applied.
        _lastInsertedComponent = _assembly.AddComponent5(
            path,
            (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
            string.Empty,
            false,
            string.Empty,
            _insertIndex * 0.2,
            0.0,
            0.0);

        if (_lastInsertedComponent == null)
            throw new Exception("Failed to insert component into assembly: " + path);

        _insertIndex++;
        _model.ClearSelection2(true);
        _lastInsertedComponent.Select2(false, 0);

        // Newly inserted components are unfixed so mate operations can move them.
        _assembly.UnfixComponent();

        // The source document no longer needs to stay open once insertion is complete.
        if (openedModel != null)
            _swApp.CloseDoc(openedModel.GetTitle());

        Console.WriteLine($"Inserted component: {path}");
        return _lastInsertedComponent;
    }

    // These helpers classify reference names so the selection code knows whether it is dealing
    // with a plane, an axis, or a named face.
    private static bool IsPlaneReference(string referenceName)
    {
        if (string.IsNullOrWhiteSpace(referenceName)) return false;
        return referenceName.IndexOf("plane", StringComparison.OrdinalIgnoreCase) >= 0
            || referenceName.IndexOf("ebene", StringComparison.OrdinalIgnoreCase) >= 0
            || referenceName.Equals("Oben", StringComparison.OrdinalIgnoreCase)
            || referenceName.Equals("Vorne", StringComparison.OrdinalIgnoreCase)
            || referenceName.Equals("Rechts", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAxisReference(string referenceName)
    {
        if (string.IsNullOrWhiteSpace(referenceName)) return false;
        return referenceName.IndexOf("axis", StringComparison.OrdinalIgnoreCase) >= 0
            || referenceName.IndexOf("achse", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    /// <summary>
    /// Shared selection helper used by all mate methods.
    ///
    /// It supports:
    /// - assembly planes
    /// - component planes
    /// - component axes
    /// - named faces inside a part
    /// </summary>
    private bool SelectReferenceGeneric(ModelDocExtension swExt, Component2 comp, string referenceName, string selectionToken, bool append)
    {
        bool looksLikePlane = IsPlaneReference(referenceName);
        bool looksLikeAxis = IsAxisReference(referenceName);

        // If the reference is not a plane or axis, we treat it as a named face and map
        // that part face into the assembly context.
        if (comp != null && !looksLikePlane && !looksLikeAxis)
        {
            PartDoc partDoc = comp.GetModelDoc2() as PartDoc;
            Entity faceInPart = partDoc?.GetEntityByName(referenceName, (int)swSelectType_e.swSelFACES) as Entity;
            Entity faceInAssembly = (faceInPart != null) ? comp.GetCorrespondingEntity(faceInPart) as Entity : null;
            if (faceInAssembly != null)
                return faceInAssembly.Select4(append, null);
        }

        string selType = looksLikePlane ? "PLANE" : (looksLikeAxis ? "AXIS" : "FACE");
        return swExt.SelectByID2(selectionToken, selType, 0, 0, 0, append, 0, null, 0);
    }

    /// <summary>
    /// Aligns a newly inserted component to the assembly origin by mating the three default planes.
    /// </summary>
    public void MateToOrigin(Component2 insertedComponent)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        if (insertedComponent == null)
            throw new InvalidOperationException("Inserted component is null.");

        ModelDoc2 compModel = (ModelDoc2)insertedComponent.GetModelDoc2();

        bool isAssembly = compModel.GetType() ==
                          (int)swDocumentTypes_e.swDocASSEMBLY;

        string asmName = _model.GetTitle().Replace(".SLDASM", "");
        string compName = insertedComponent.Name2;

        ModelDocExtension swExt = _model.Extension;

        string[] assemblyPlanes = { "Oben", "Vorne", "Rechts" };
        string[] partPlanes     = { "Ebene oben", "Ebene vorne", "Ebene rechts" };

        for (int i = 0; i < 3; i++)
        {
            _model.ClearSelection2(true);

            // If the inserted component is itself an assembly, its default planes use assembly names.
            // If it is a part, its default planes use part plane names.
            string compPlaneName = isAssembly
                ? assemblyPlanes[i]
                : partPlanes[i];

            string compSelection = isAssembly
                ? $"{compPlaneName}@{compName}@{asmName}"
                : $"{compPlaneName}@{compName}@{asmName}";

            string asmSelection = assemblyPlanes[i];

            bool s1 = swExt.SelectByID2(
                asmSelection,
                "PLANE",
                0, 0, 0,
                false,
                0,
                null,
                0);

            bool s2 = swExt.SelectByID2(
                compSelection,
                "PLANE",
                0, 0, 0,
                true,
                0,
                null,
                0);

            if (s1 && s2)
            {
                int mateError = 0;

                _assembly.AddMate5(
                    (int)swMateType_e.swMateCOINCIDENT,
                    (int)swMateAlign_e.swMateAlignALIGNED,
                    false,
                    0, 0, 0, 0, 0, 0, 0, 0,
                    false,
                    false,
                    0,
                    out mateError);
            }
        }

        _model.EditRebuild3();
    }
    /// <summary>
    /// Creates either a coincident mate or a distance mate between two references.
    ///
    /// If offset is zero, we create a true coincident mate.
    /// If offset is non-zero, we create a distance mate using that offset.
    /// </summary>
    public void MateCoincident(Component2 comp1, string ref1, Component2 comp2, string ref2, double offset = 0, bool flipAlignment = false)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        // Build the top-level assembly title used by SolidWorks selection strings.
        string asmName = _model.GetTitle().Split('.')[0];

        // If a component is null, the reference belongs to the top-level assembly itself.
        string selection1 = (comp1 != null) ? $"{ref1}@{comp1.Name2}@{asmName}" : ref1;
        string selection2 = (comp2 != null) ? $"{ref2}@{comp2.Name2}@{asmName}" : ref2;

        ModelDocExtension swExt = _model.Extension;
        _model.ClearSelection2(true);

        bool s1 = SelectReferenceGeneric(swExt, comp1, ref1, selection1, false);
        bool s2 = SelectReferenceGeneric(swExt, comp2, ref2, selection2, true);

        if (s1 && s2)
        {
            int mateError = 0;
        
            // Zero offset means true coincidence.
            // Non-zero offset means a distance mate with a fixed spacing.
            swMateType_e mateType = (offset == 0) 
                ? swMateType_e.swMateCOINCIDENT 
                : swMateType_e.swMateDISTANCE;

            _assembly.AddMate5(
                (int)mateType,
                flipAlignment ? (int)swMateAlign_e.swMateAlignANTI_ALIGNED : (int)swMateAlign_e.swMateAlignALIGNED,
                flipAlignment,
                offset,
                offset,
                offset,
                0, 0, 0, 0, 0,
                false,
                false,
                0,
                out mateError);

            if (mateError > 1)
                Console.WriteLine($"Mate failed with error code: {mateError}");
        }
        else
        {
            Console.WriteLine($"Selection failed for {selection1} or {selection2}");
        }

        _model.EditRebuild3();
    }

    /// <summary>
    /// Creates an angle mate between two references.
    /// </summary>
    public void MateAngle(Component2 comp1, string ref1, Component2 comp2, string ref2, double angleDeg, bool flipAlignment = false)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        string asmName = _model.GetTitle().Split('.')[0];
        string selection1 = (comp1 != null) ? $"{ref1}@{comp1.Name2}@{asmName}" : ref1;
        string selection2 = (comp2 != null) ? $"{ref2}@{comp2.Name2}@{asmName}" : ref2;

        ModelDocExtension swExt = _model.Extension;
        _model.ClearSelection2(true);

        bool s1 = SelectReferenceGeneric(swExt, comp1, ref1, selection1, false);
        bool s2 = SelectReferenceGeneric(swExt, comp2, ref2, selection2, true);

        if (s1 && s2)
        {
            int mateError = 0;
            double angleRadians = angleDeg * Math.PI / 180.0;

            _assembly.AddMate5(
                (int)swMateType_e.swMateANGLE,
                flipAlignment ? (int)swMateAlign_e.swMateAlignANTI_ALIGNED : (int)swMateAlign_e.swMateAlignALIGNED,
                flipAlignment,
                0.0,
                0.0,
                0.0,
                0.0, 0.0, angleRadians, angleRadians, angleRadians,
                false,
                false,
                0,
                out mateError);

            if (mateError > 1)
                Console.WriteLine($"Angle mate failed with error code: {mateError}");
        }
        else
        {
            Console.WriteLine($"Selection failed for angle mate: {selection1} or {selection2}");
        }

        _model.EditRebuild3();
    }

    /// <summary>
    /// Creates a parallel mate between two references, usually axes or planes.
    /// </summary>
    public void MateParallel(Component2 comp1, string ref1, Component2 comp2, string ref2)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        string asmName = _model.GetTitle().Split('.')[0];
        string selection1 = (comp1 != null) ? $"{ref1}@{comp1.Name2}@{asmName}" : ref1;
        string selection2 = (comp2 != null) ? $"{ref2}@{comp2.Name2}@{asmName}" : ref2;

        ModelDocExtension swExt = _model.Extension;
        _model.ClearSelection2(true);

        bool s1 = SelectReferenceGeneric(swExt, comp1, ref1, selection1, false);
        bool s2 = SelectReferenceGeneric(swExt, comp2, ref2, selection2, true);

        if (s1 && s2)
        {
            int mateError = 0;

            _assembly.AddMate5(
                (int)swMateType_e.swMatePARALLEL,
                (int)swMateAlign_e.swMateAlignALIGNED,
                false,
                0, 0, 0, 0, 0, 0, 0, 0,
                false,
                false,
                0,
                out mateError);

            if (mateError > 1)
                Console.WriteLine($"Parallel mate failed for {selection1}/{selection2}. Error: {mateError}");
        }
        else
        {
            Console.WriteLine($"Selection failed for Parallel mate: {selection1} or {selection2}");
        }

        _model.EditRebuild3();
    }

    /// <summary>
    /// Convenience helper that aligns the X, Y, and Z axes between two components.
    /// </summary>
    public void MateAxes(Component2 comp1, Component2 comp2)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        string[] axes = { "X-Achse", "Y-Achse", "Z-Achse" };
        foreach (string axis in axes)
        {
            MateCoincident(comp1, axis, comp2, axis);
        }
    }

    /// <summary>
    /// Creates a linear component pattern using one reference direction and one or more seed components.
    /// </summary>
    public Feature LinearPattern(Component2 directionOwner, string directionRef, int instanceCount, double spacing, params Component2[] seedComponents)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        if (seedComponents == null || seedComponents.Length == 0)
            throw new ArgumentException("At least one seed component is required.", nameof(seedComponents));

        if (instanceCount < 2)
            throw new ArgumentOutOfRangeException(nameof(instanceCount), "Linear pattern count must be at least 2.");

    string asmName = _model.GetTitle().Split('.')[0];
    string directionSelection = (directionOwner != null) ? $"{directionRef}@{directionOwner.Name2}@{asmName}" : directionRef;
    ModelDocExtension swExt = _model.Extension;
    SelectionMgr selectionMgr = _model.SelectionManager as SelectionMgr;
    if (selectionMgr == null)
        throw new Exception("Could not access selection manager for linear component pattern");

    // SolidWorks can be very sensitive to selection marks in component patterns.
    // We sweep through the common working mark combinations until one is accepted.
    foreach (int directionMark in new[] { 2, 1, 4, 256 })
    {
        foreach (int componentMark in new[] { 1, 4, 256, 2 })
        {
            foreach (bool directionFirst in new[] { true, false })
            {
                SelectData directionSelectData = selectionMgr.CreateSelectData();
                SelectData componentSelectData = selectionMgr.CreateSelectData();
                if (directionSelectData == null || componentSelectData == null)
                    continue;

                directionSelectData.Mark = directionMark;
                componentSelectData.Mark = componentMark;

                _model.ClearSelection2(true);

                bool directionSelected;
                int selectedSeedCount;

                if (directionFirst)
                {
                    directionSelected = SelectPatternDirection(false, directionSelectData);
                    selectedSeedCount = SelectSeedComponents(true, componentSelectData);
                }
                else
                {
                    selectedSeedCount = SelectSeedComponents(false, componentSelectData);
                    directionSelected = SelectPatternDirection(true, directionSelectData);
                }

                if (!directionSelected || selectedSeedCount != seedComponents.Length)
                    continue;

                Feature patternFeature = _model.FeatureManager.FeatureLinearPattern5(
                    instanceCount,
                    spacing,
                    1,
                    0.0,
                    false,
                    false,
                    string.Empty,
                    string.Empty,
                    false,
                    false,
                    false,
                    false,
                    true,
                    true,
                    false,
                    false,
                    false,
                    false,
                    0.0,
                    0.0,
                    false,
                    false) as Feature;

                if (patternFeature != null)
                {
                    _model.EditRebuild3();
                    return patternFeature;
                }
            }
        }
    }

        throw new Exception("Failed to create linear component pattern");

    bool SelectPatternDirection(bool append, SelectData directionSelectData)
    {
        if (directionOwner != null && !IsPlaneReference(directionRef) && !IsAxisReference(directionRef))
        {
            PartDoc partDoc = directionOwner.GetModelDoc2() as PartDoc;
            Entity faceInPart = partDoc?.GetEntityByName(directionRef, (int)swSelectType_e.swSelFACES) as Entity;
            Entity faceInAssembly = (faceInPart != null) ? directionOwner.GetCorrespondingEntity(faceInPart) as Entity : null;
            return faceInAssembly != null && faceInAssembly.Select4(append, directionSelectData);
        }

        string selType = IsPlaneReference(directionRef) ? "PLANE" : (IsAxisReference(directionRef) ? "AXIS" : "FACE");
        return swExt.SelectByID2(directionSelection, selType, 0, 0, 0, append, directionSelectData.Mark, null, 0);
    }

    int SelectSeedComponents(bool appendFirst, SelectData componentSelectData)
    {
        int count = 0;
        bool append = appendFirst;
        foreach (Component2 seedComponent in seedComponents)
        {
            if (seedComponent != null && seedComponent.Select4(append, componentSelectData, false))
            {
                count++;
                append = true;
            }
        }

        return count;
    }
    }

    /// <summary>
    /// Creates a circular component pattern around one reference axis.
    /// </summary>
    public Feature CircularPattern(Component2 axisOwner, string axisRef, int instanceCount, double angle, params Component2[] seedComponents)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        if (seedComponents == null || seedComponents.Length == 0)
            throw new ArgumentException("At least one seed component is required.", nameof(seedComponents));

        if (instanceCount < 2)
            throw new ArgumentOutOfRangeException(nameof(instanceCount), "Circular pattern count must be at least 2.");

    string asmName = _model.GetTitle().Split('.')[0];
    string axisSelection = (axisOwner != null) ? $"{axisRef}@{axisOwner.Name2}@{asmName}" : axisRef;
    ModelDocExtension swExt = _model.Extension;
    SelectionMgr selectionMgr = _model.SelectionManager as SelectionMgr;
    if (selectionMgr == null)
        throw new Exception("Could not access selection manager for circular component pattern");

    // Circular patterns use the same idea as the linear pattern above:
    // cycle through the common selection-mark combinations until SolidWorks accepts one.
    foreach (int axisMark in new[] { 1, 2, 4, 256 })
    {
        foreach (int componentMark in new[] { 1, 4, 256, 2 })
        {
            foreach (bool axisFirst in new[] { true, false })
            {
                SelectData axisSelectData = selectionMgr.CreateSelectData();
                SelectData componentSelectData = selectionMgr.CreateSelectData();
                if (axisSelectData == null || componentSelectData == null)
                    continue;

                axisSelectData.Mark = axisMark;
                componentSelectData.Mark = componentMark;

                _model.ClearSelection2(true);

                bool axisSelected;
                int selectedSeedCount;

                if (axisFirst)
                {
                    axisSelected = SelectPatternAxis(false, axisSelectData);
                    selectedSeedCount = SelectSeedComponents(true, componentSelectData);
                }
                else
                {
                    selectedSeedCount = SelectSeedComponents(false, componentSelectData);
                    axisSelected = SelectPatternAxis(true, axisSelectData);
                }

                if (!axisSelected || selectedSeedCount != seedComponents.Length)
                    continue;

                Feature patternFeature = _model.FeatureManager.FeatureCircularPattern5(
                    instanceCount,
                    angle,
                    false,
                    string.Empty,
                    false,
                    true,
                    false,
                    false,
                    false,
                    false,
                    0,
                    0.0,
                    string.Empty,
                    false) as Feature;

                if (patternFeature != null)
                {
                    _model.EditRebuild3();
                    return patternFeature;
                }
            }
        }
    }

        throw new Exception("Failed to create circular component pattern");

    bool SelectPatternAxis(bool append, SelectData axisSelectData)
    {
        if (axisOwner != null && !IsPlaneReference(axisRef) && !IsAxisReference(axisRef))
        {
            PartDoc partDoc = axisOwner.GetModelDoc2() as PartDoc;
            Entity faceInPart = partDoc?.GetEntityByName(axisRef, (int)swSelectType_e.swSelFACES) as Entity;
            Entity faceInAssembly = (faceInPart != null) ? axisOwner.GetCorrespondingEntity(faceInPart) as Entity : null;
            return faceInAssembly != null && faceInAssembly.Select4(append, axisSelectData);
        }

        string selType = IsPlaneReference(axisRef) ? "PLANE" : (IsAxisReference(axisRef) ? "AXIS" : "FACE");
        return swExt.SelectByID2(axisSelection, selType, 0, 0, 0, append, axisSelectData.Mark, null, 0);
    }

    int SelectSeedComponents(bool appendFirst, SelectData componentSelectData)
    {
        int count = 0;
        bool append = appendFirst;
        foreach (Component2 seedComponent in seedComponents)
        {
            if (seedComponent != null && seedComponent.Select4(append, componentSelectData, false))
            {
                count++;
                append = true;
            }
        }

        return count;
    }
    }
}

