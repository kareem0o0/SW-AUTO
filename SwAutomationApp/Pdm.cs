using System;
using System.IO;
using EPDM.Interop.epdm;
using SolidWorks.Interop.sldworks; // Added this
using SolidWorks.Interop.swconst;  // Added this

namespace SwAutomation.Pdm
{
    public sealed class PdmModule
    {
        private IEdmVault5 _vault;
        private const string VaultRoot = @"C:\Users\kareem.salah\PDM\Birr Machines PDM";

        public void Login()
        {
            _vault = new EdmVault5();
            _vault.Login("kareem.salah", "976431852@KEmo", "Birr Machines PDM");

            if (!_vault.IsLoggedIn)
                throw new Exception("PDM Login failed.");

            Console.WriteLine("Logged into vault");
        }

        public void AddExistingFileToPdm(string localFilePath, string subFolder = "60_Tests")
        {
            var vault11 = (IEdmVault11)_vault;
            var snGen = (IEdmSerNoGen7)vault11.CreateUtility(EdmUtility.EdmUtil_SerNoGen);
            snGen.GetSerialNumberNames(out string[] snNames);
            if (snNames?.Length == 0) throw new Exception("No Serial Numbers found.");
            var snValueObj = snGen.AllocSerNoValue(snNames[0], 0, "", 0, 0, 0, 0);
            string snString = snValueObj.Value;

            string targetFolderPath = Path.Combine(VaultRoot, subFolder);
            string destinationPath = Path.Combine(targetFolderPath, snString + Path.GetExtension(localFilePath));

            if (!Directory.Exists(targetFolderPath)) Directory.CreateDirectory(targetFolderPath);

            File.Copy(localFilePath, destinationPath, true);

            IEdmFolder5 folder = _vault.GetFolderFromPath(targetFolderPath);
            folder.AddFile(0, destinationPath, "", 1);

            Console.WriteLine($"Vaulted as: {Path.GetFileName(destinationPath)}");
        }

        public string SaveAsPdm(ModelDoc2 swModel, string subFolder = "60_Tests")
{
        // 1. Generate the PDM Serial Number string in this same method
        var vault11 = (IEdmVault11)_vault;
        var snGen = (IEdmSerNoGen7)vault11.CreateUtility(EdmUtility.EdmUtil_SerNoGen);
        snGen.GetSerialNumberNames(out string[] snNames);
        if (snNames?.Length == 0) throw new Exception("No Serial Numbers found.");
        var snValueObj = snGen.AllocSerNoValue(snNames[0], 0, "", 0, 0, 0, 0);
        string snString = snValueObj.Value;

        // 2. Determine file extension based on Document Type
        string extension = ".sldprt";
        int type = swModel.GetType();
        if (type == (int)swDocumentTypes_e.swDocASSEMBLY) extension = ".sldasm";
        else if (type == (int)swDocumentTypes_e.swDocDRAWING) extension = ".slddrw";

        // 3. Build the full path inside the Vault Root
        string targetDir = Path.Combine(VaultRoot, subFolder);
        string fullPath = Path.Combine(targetDir, snString + extension);

        // 4. Ensure the physical folder exists in the Windows File System
        if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

        // 5. Save directly from SolidWorks to PDM using your 3-argument style
        // (int)1 = swSaveAsOptions_Silent
        // (int)0 = swSaveAsCurrentVersion
        swModel.SaveAs3(fullPath, 0, 1);

        // 6. Register the file in the PDM Database (The "Add" step)
        IEdmFolder5 folder = _vault.GetFolderFromPath(targetDir);
        if (folder == null) throw new Exception("Vault folder object not found.");
        
        // AddFile(ParentWnd, Path, Name, Flags) -> 1 is EdmAdd_Simple
        folder.AddFile(0, fullPath, "", 1);

        Console.WriteLine($"Saved and Vaulted: {Path.GetFileName(fullPath)}");
        return fullPath;
}
    }
}
