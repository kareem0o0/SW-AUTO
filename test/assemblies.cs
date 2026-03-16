using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation.TestHarness;

internal sealed class Assembly
{
    private readonly SldWorks _swApp;
    private ModelDoc2? _model;
    private AssemblyDoc? _assembly;
    private string _folder = string.Empty;
    private int _insertIndex;

    public Assembly(SldWorks swApp)
    {
        _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
    }

    public string CreateAssembly(string outFolder, string fileName = "Assembly.SLDASM", bool closeAfterCreate = false)
    {
        Directory.CreateDirectory(outFolder);

        string template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
        ModelDoc2 model = (ModelDoc2)_swApp.NewDocument(template, 0, 0, 0);
        string fullPath = Path.Combine(outFolder, fileName);

        model.SaveAs3(fullPath, 0, 1);
        Console.WriteLine($"Assembly saved to: {fullPath}");
        _insertIndex = 0;

        if (closeAfterCreate)
        {
            _swApp.CloseDoc(model.GetTitle());
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

        string path = Path.GetFullPath(
            Path.IsPathRooted(componentPath)
                ? componentPath
                : Path.Combine(_folder, componentPath));

        int docType = Path.GetExtension(path).Equals(".sldasm", StringComparison.OrdinalIgnoreCase)
            ? (int)swDocumentTypes_e.swDocASSEMBLY
            : (int)swDocumentTypes_e.swDocPART;

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

        Component2 component = _assembly.AddComponent5(
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
        component.Select2(false, 0);
        _assembly.UnfixComponent();

        if (openedModel != null)
            _swApp.CloseDoc(openedModel.GetTitle());

        Console.WriteLine($"Inserted component: {path}");
        return component;
    }

    public void MatePlans(Component2 insertedComponent)
    {
        if (_model == null || _assembly == null)
            throw new InvalidOperationException("No open assembly.");

        string assemblyName = _model.GetTitle().Replace(".SLDASM", "");
        string componentName = insertedComponent.Name2;
        string[] assemblyPlanes = { "Oben", "Vorne", "Rechts" };

        for (int i = 0; i < assemblyPlanes.Length; i++)
        {
            _model.ClearSelection2(true);

            bool selectAssemblyPlane = _model.Extension.SelectByID2(
                assemblyPlanes[i],
                "PLANE",
                0,
                0,
                0,
                false,
                0,
                null,
                0);

            bool selectComponentPlane = _model.Extension.SelectByID2(
                $"{assemblyPlanes[i]}@{componentName}@{assemblyName}",
                "PLANE",
                0,
                0,
                0,
                true,
                0,
                null,
                0);

            if (!selectAssemblyPlane || !selectComponentPlane)
                continue;

            int mateError = 0;
            _assembly.AddMate5(
                (int)swMateType_e.swMateCOINCIDENT,
                (int)swMateAlign_e.swMateAlignALIGNED,
                false,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                false,
                false,
                0,
                out mateError);
        }

        _model.EditRebuild3();
    }
}
