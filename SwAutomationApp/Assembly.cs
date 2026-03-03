using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation;

public sealed class Assembly
{
    private readonly SldWorks _swApp;
    private ModelDoc2 _model = null;
    private AssemblyDoc _assembly = null;
    private string _folder = string.Empty;
    private int _insertIndex = 0;
    private Component2 swComponent = null;
    public Assembly(SldWorks swApp) => _swApp = swApp;

    public void CreateAssembly(string outFolder, string fileName = "Assembly.SLDASM", bool closeAfterCreate = false)
    {
        Directory.CreateDirectory(outFolder);
        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
        ModelDoc2 model = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
        string fullPath = Path.Combine(outFolder, fileName);
        model.SaveAs3(fullPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent);

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
            return;
        }

        _model = model;
        _assembly = model as AssemblyDoc;
        _folder = outFolder;

        int activateErrors = 0;
        _swApp.ActivateDoc3(_model.GetTitle(), false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref activateErrors);
        Console.WriteLine("Assembly remains open for further operations.");
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

    _insertIndex++;
    _model.ClearSelection2(true);
    swComponent.Select2(false, 0);

    _assembly.UnfixComponent();
    

    if (openedModel != null) _swApp.CloseDoc(openedModel.GetTitle());

    Console.WriteLine($"Inserted component: {path}");
    return swComponent;
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

}
