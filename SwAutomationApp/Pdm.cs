using System;
using System.Collections.Generic;
using System.IO;
using EPDM.Interop.epdm;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation.Pdm
{
    public sealed class BirrDataCardValues
    {
        // These properties map to the Birr PDM data-card fields.
        public string DrawingNumber { get; set; } = string.Empty;
        public string Title { get; set; } = "Stator Sheet Metal";
        public string Subtitle { get; set; } = "Automated Generation";
        public string Project { get; set; } = "665_Birr_Project";
        public string Customer { get; set; } = "Birr Machines AG";
        public string CustomerOrder { get; set; } = "66";
        public string Type { get; set; } = "44";
        public string Unit { get; set; } = "kg";
        public string CreatedFrom { get; set; } = "22";
        public string ReplacementFor { get; set; } = "11";
        public string DataCheck { get; set; } = "X";

        public static BirrDataCardValues CreateDefault()
        {
            return new BirrDataCardValues();
        }

        public Dictionary<string, string> ToDictionary(string generatedDrawingNumber = null)
        {
            string drawingNumber = string.IsNullOrWhiteSpace(DrawingNumber)
                ? (generatedDrawingNumber ?? string.Empty)
                : DrawingNumber;

            return new Dictionary<string, string>
            {
                { "Zeichnungs-Nr.:", drawingNumber },
                { "Title", Title ?? string.Empty },
                { "Subtitle", Subtitle ?? string.Empty },
                { "Projekt", Project ?? string.Empty },
                { "Kunde", Customer ?? string.Empty },
                { "Kundenauftrag", CustomerOrder ?? string.Empty },
                { "Type", Type ?? string.Empty },
                { "Einheit", Unit ?? string.Empty },
                { "Entstand aus", CreatedFrom ?? string.Empty },
                { "Ersatz f\u00FCr", ReplacementFor ?? string.Empty },
                { "DataCheck", DataCheck ?? string.Empty }
            };
        }
    }

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
            IEdmVault11 vault11 = (IEdmVault11)_vault;
            IEdmSerNoGen7 snGen = (IEdmSerNoGen7)vault11.CreateUtility(EdmUtility.EdmUtil_SerNoGen);
            snGen.GetSerialNumberNames(out string[] snNames);
            if (snNames?.Length == 0)
                throw new Exception("No Serial Numbers found.");

            var snValueObj = snGen.AllocSerNoValue(snNames[0], 0, string.Empty, 0, 0, 0, 0);
            string snString = snValueObj.Value;

            string targetFolderPath = Path.Combine(VaultRoot, subFolder);
            string destinationPath = Path.Combine(targetFolderPath, snString + Path.GetExtension(localFilePath));

            if (!Directory.Exists(targetFolderPath))
                Directory.CreateDirectory(targetFolderPath);

            File.Copy(localFilePath, destinationPath, true);

            IEdmFolder5 folder = _vault.GetFolderFromPath(targetFolderPath);
            folder.AddFile(0, destinationPath, string.Empty, 1);

            Console.WriteLine($"Vaulted as: {Path.GetFileName(destinationPath)}");
        }

        public string SaveAsPdm(ModelDoc2 swModel, string subFolder = "60_Tests", BirrDataCardValues dataCardValues = null)
        {
            IEdmVault11 vault11 = (IEdmVault11)_vault;
            IEdmSerNoGen7 snGen = (IEdmSerNoGen7)vault11.CreateUtility(EdmUtility.EdmUtil_SerNoGen);
            snGen.GetSerialNumberNames(out string[] snNames);
            if (snNames?.Length == 0)
                throw new Exception("No Serial Numbers found.");

            var snValueObj = snGen.AllocSerNoValue(snNames[0], 0, string.Empty, 0, 0, 0, 0);
            string snString = snValueObj.Value;

            string extension = ".sldprt";
            int type = swModel.GetType();
            if (type == (int)swDocumentTypes_e.swDocASSEMBLY)
                extension = ".sldasm";
            else if (type == (int)swDocumentTypes_e.swDocDRAWING)
                extension = ".slddrw";

            string targetDir = Path.Combine(VaultRoot, subFolder);
            string fullPath = Path.Combine(targetDir, snString + extension);

            if (!Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            swModel.SaveAs3(fullPath, 0, 1);

            IEdmFolder5 folder = _vault.GetFolderFromPath(targetDir);
            if (folder == null)
                throw new Exception("Vault folder object not found.");

            folder.AddFile(0, fullPath, string.Empty, 1);

            if (dataCardValues != null)
            {
                UpdateBirrDataCard(fullPath, dataCardValues.ToDictionary(snString));
            }

            Console.WriteLine($"Saved and Vaulted: {Path.GetFileName(fullPath)}");
            return fullPath;
        }

        public void GetDataCardValues(string filePath)
        {
            string fullPath = Path.IsPathRooted(filePath)
                ? filePath
                : Path.Combine(VaultRoot, filePath);
            IEdmFile5 file = _vault.GetFileFromPath(fullPath, out IEdmFolder5 parentFolder);

            if (file == null)
            {
                Console.WriteLine("File not found.");
                return;
            }

            IEdmEnumeratorVariable5 varEnum = (IEdmEnumeratorVariable5)file.GetEnumeratorVariable();
            IEdmVariableMgr5 varMgr = (IEdmVariableMgr5)_vault;

            IEdmStrLst5 cfgList = file.GetConfigurations();
            IEdmPos5 cfgPos = cfgList.GetHeadPosition();

            while (!cfgPos.IsNull)
            {
                string cfgName = cfgList.GetNext(cfgPos);
                Console.WriteLine($"\n--- Configuration: [{cfgName}] ---");

                IEdmPos5 varPos = varMgr.GetFirstVariablePosition();
                while (!varPos.IsNull)
                {
                    IEdmVariable5 variable = varMgr.GetNextVariable(varPos);
                    varEnum.GetVar(variable.Name, cfgName, out object varValue);

                    if (varValue != null && !string.IsNullOrEmpty(varValue.ToString()))
                    {
                        Console.WriteLine($"{variable.Name}: {varValue}");
                    }
                }
            }
        }

        public void UpdateBirrDataCard(string relativePath, Dictionary<string, string> values)
        {
            string fullPath = Path.IsPathRooted(relativePath)
                ? relativePath
                : Path.Combine(VaultRoot, relativePath);
            IEdmFile5 file = _vault.GetFileFromPath(fullPath, out IEdmFolder5 parentFolder);
            if (file == null)
                throw new Exception($"File not found: {fullPath}");

            if (!file.IsLocked)
                file.LockFile(parentFolder.ID, 0);

            IEdmEnumeratorVariable5 varEnum = (IEdmEnumeratorVariable5)file.GetEnumeratorVariable();
            IEdmVariableMgr5 varMgr = (IEdmVariableMgr5)_vault;

            HashSet<string> configsToUpdate = new(StringComparer.OrdinalIgnoreCase) { "@" };
            IEdmStrLst5 cfgList = file.GetConfigurations();
            if (cfgList != null)
            {
                IEdmPos5 cfgPos = cfgList.GetHeadPosition();
                while (!cfgPos.IsNull)
                {
                    string cfgName = cfgList.GetNext(cfgPos);
                    if (!string.IsNullOrWhiteSpace(cfgName))
                    {
                        configsToUpdate.Add(cfgName);
                    }
                }
            }

            foreach (string config in configsToUpdate)
            {
                foreach (KeyValuePair<string, string> entry in values)
                {
                    try
                    {
                        string actualVarName = null;
                        IEdmPos5 pos = varMgr.GetFirstVariablePosition();
                        while (!pos.IsNull)
                        {
                            IEdmVariable5 variable = varMgr.GetNextVariable(pos);
                            if (variable.Name.Equals(entry.Key, StringComparison.OrdinalIgnoreCase) ||
                                variable.Name.StartsWith(entry.Key.Replace(":", string.Empty), StringComparison.OrdinalIgnoreCase))
                            {
                                actualVarName = variable.Name;
                                break;
                            }
                        }

                        if (!string.IsNullOrEmpty(actualVarName))
                        {
                            varEnum.SetVar(actualVarName, config, entry.Value ?? string.Empty, false);
                        }
                        else
                        {
                            Console.WriteLine($"SKIPPED: Variable matching '{entry.Key}' not found in Vault.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error setting {entry.Key}: {ex.Message}");
                    }
                }
            }

            ((IEdmEnumeratorVariable8)varEnum).CloseFile(true);
            file.UnlockFile(0, "Automated data card sync");
        }

        public void FillBirrDataCard(string relativePath)
        {
            BirrDataCardValues birrData = BirrDataCardValues.CreateDefault();
            UpdateBirrDataCard(relativePath, birrData.ToDictionary());
            Console.WriteLine($"Full Data Card populated for: {relativePath}");
        }
    }
}
