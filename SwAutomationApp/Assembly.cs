using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwAutomation.Pdm;

namespace SwAutomation;

public sealed class Assembly
{
    private readonly SldWorks _swApp;
    private readonly PdmModule _pdm;
    private ModelDoc2 _model = null;
    private AssemblyDoc _assembly = null;
    private string _folder = string.Empty;
    private int _insertIndex = 0;
    private Component2 swComponent = null;
    public Assembly(SldWorks swApp, PdmModule pdm)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        _pdm = pdm ?? throw new ArgumentNullException(nameof(pdm));
    }

    public string CreateAssembly(string outFolder, string fileName = "Assembly.SLDASM", bool closeAfterCreate = false, bool SaveToPdm = false)
    {
        Directory.CreateDirectory(outFolder);
        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
        ModelDoc2 model = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
        string fullPath;
        if (SaveToPdm)
        {
            fullPath = _pdm.SaveAsPdm(model, outFolder);
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
            _swApp.CloseDoc(model.GetTitle());
            _model = null;
            _assembly = null;
            _folder = string.Empty;
            swComponent = null;
            Console.WriteLine("Assembly closed after creating.");
            return Path.GetFileName(fullPath);
        }

        _model = model;
        _assembly = model as AssemblyDoc;
        _folder = Path.GetDirectoryName(fullPath) ?? string.Empty;

        int activateErrors = 0;
        _swApp.ActivateDoc3(_model.GetTitle(), false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref activateErrors);
        Console.WriteLine("Assembly remains open for further operations.");
        return Path.GetFileName(fullPath);
    }

    public Component2 InsertComponentToOpenAssembly(string componentPath)
    {
        
        
    if (_model == null || _assembly == null)
        throw new InvalidOperationException("No open assembly.");

    if (string.IsNullOrWhiteSpace(componentPath))
        throw new ArgumentException("Component path is required.", nameof(componentPath));

    // Resolve full path
    string path = Path.GetFullPath(
        Path.IsPathRooted(componentPath)
            ? componentPath
            : Path.Combine(_folder, componentPath));

    // Determine document type
    int docType = Path.GetExtension(path)
        .Equals(".sldasm", StringComparison.OrdinalIgnoreCase)
            ? (int)swDocumentTypes_e.swDocASSEMBLY
            : (int)swDocumentTypes_e.swDocPART;

    // Open and activate
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

    _swApp.ActivateDoc3(
        _model.GetTitle(),
        false,
        (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
        ref activateErrors);

    // Insert component
    swComponent = _assembly.AddComponent5(
        path,
        (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig,
        string.Empty,
        false,
        string.Empty,
        _insertIndex * 0.2,
        0.0,
        0.0);

    if (swComponent == null)
        throw new Exception("Failed to insert component into assembly: " + path);

    _insertIndex++;
    _model.ClearSelection2(true);
    swComponent.Select2(false, 0);

    _assembly.UnfixComponent();
    

    if (openedModel != null) _swApp.CloseDoc(openedModel.GetTitle());

    Console.WriteLine($"Inserted component: {path}");
    return swComponent;
}

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

private bool SelectReferenceGeneric(ModelDocExtension swExt, Component2 comp, string referenceName, string selectionToken, bool append)
{
    bool looksLikePlane = IsPlaneReference(referenceName);
    bool looksLikeAxis = IsAxisReference(referenceName);

    // Generic named-face resolution for all mate methods.
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

public void mate_plans(Component2 insertedComponent)
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

        string compPlaneName = isAssembly
            ? assemblyPlanes[i]      // assembly-to-assembly
            : partPlanes[i];         // assembly-to-part

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
public void ApplyCoincedentMate(Component2 comp1, string ref1, Component2 comp2, string ref2, double offset = 0, bool flipAlignment = false)
{
    if (_model == null || _assembly == null)
        throw new InvalidOperationException("No open assembly.");

    // 1. Get the top-level assembly name (needed for the selection string)
    // Remove the extension to get the clean title
    string asmName = _model.GetTitle().Split('.')[0];

    // 2. Build Selection Strings
    // Format: "PlaneName@InstanceName@AssemblyName"
    // If a component is null, we assume the reference belongs to the top-level assembly itself
    string selection1 = (comp1 != null) ? $"{ref1}@{comp1.Name2}@{asmName}" : ref1;
    string selection2 = (comp2 != null) ? $"{ref2}@{comp2.Name2}@{asmName}" : ref2;

    ModelDocExtension swExt = _model.Extension;
    _model.ClearSelection2(true);

    bool s1 = SelectReferenceGeneric(swExt, comp1, ref1, selection1, false);
    bool s2 = SelectReferenceGeneric(swExt, comp2, ref2, selection2, true);

    if (s1 && s2)
    {
        int mateError = 0;
        
        // Decide Mate Type: If offset is 0, use Coincident. Otherwise, use Distance.
        swMateType_e mateType = (offset == 0) 
            ? swMateType_e.swMateCOINCIDENT 
            : swMateType_e.swMateDISTANCE;

        // AddMate5 Parameters:
        // Type, Alignment, Flip, Distance, Max, Min, etc.
        _assembly.AddMate5(
            (int)mateType,
            flipAlignment ? (int)swMateAlign_e.swMateAlignANTI_ALIGNED : (int)swMateAlign_e.swMateAlignALIGNED,
            flipAlignment,
            offset, // Distance/Offset
            offset, // Max Distance
            offset, // Min Distance
            0, 0, 0, 0, 0,
            false,  // Is SmartMate
            false,  // Use for positioning only
            0,      // Lock rotation
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
public void ApplyAngleMate(Component2 comp1, string ref1, Component2 comp2, string ref2, double angleDeg, bool flipAlignment = false)
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
public void ApplyParallelMate(Component2 comp1, string ref1, Component2 comp2, string ref2)
{
    if (_model == null || _assembly == null)
        throw new InvalidOperationException("No open assembly.");

    // 1. Get the top-level assembly name for the selection string
    string asmName = _model.GetTitle().Split('.')[0];

    // 2. Build Selection Strings: "Feature@Component@Assembly"
    string selection1 = (comp1 != null) ? $"{ref1}@{comp1.Name2}@{asmName}" : ref1;
    string selection2 = (comp2 != null) ? $"{ref2}@{comp2.Name2}@{asmName}" : ref2;

    ModelDocExtension swExt = _model.Extension;
    _model.ClearSelection2(true);

    bool s1 = SelectReferenceGeneric(swExt, comp1, ref1, selection1, false);
    bool s2 = SelectReferenceGeneric(swExt, comp2, ref2, selection2, true);

    if (s1 && s2)
    {
        int mateError = 0;

        // AddMate5 for Parallel:
        // (int)swMateType_e.swMatePARALLEL = 1
        _assembly.AddMate5(
            (int)swMateType_e.swMatePARALLEL,
            (int)swMateAlign_e.swMateAlignALIGNED,
            false, // Flip alignment
            0, 0, 0, 0, 0, 0, 0, 0, // No offsets/limits for parallel
            false, // Is SmartMate
            false, // Position only
            0,     // Lock rotation
            out mateError);

        if (mateError > 1)
            Console.WriteLine($"Parallel mate failed for {selection1}/{selection2}. Error: {mateError}");
    }
    else
    {
        Console.WriteLine($"Selection failed for Parallel mate: {selection1} or {selection2}");
    }

    _model.EditRebuild3(); // Rebuild to apply changes
}

public void ApplyCoincedentAxisMate(Component2 comp1, Component2 comp2)
{
    if (_model == null || _assembly == null)
        throw new InvalidOperationException("No open assembly.");

    string[] axes = { "X-Achse", "Y-Achse", "Z-Achse" };
    foreach (string axis in axes)
    {
        ApplyCoincedentMate(comp1, axis, comp2, axis);
    }
}

public Feature CreateLinearComponentPattern(Component2 directionOwner, string directionRef, int instanceCount, double spacing, params Component2[] seedComponents)
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

public Feature CreateCircularComponentPattern(Component2 axisOwner, string axisRef, int instanceCount, double angle, params Component2[] seedComponents)
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
