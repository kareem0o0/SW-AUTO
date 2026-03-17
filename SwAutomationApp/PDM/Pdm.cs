using System;
using System.Collections.Generic;
using System.IO;
using EPDM.Interop.epdm;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SwAutomation.Pdm
{
    /// <summary>
    /// Plain object that stores the Birr PDM data-card values.
    ///
    /// The important design idea here is:
    /// the part or assembly object owns these values,
    /// and the PDM module only reads and writes them.
    ///
    /// That keeps the data easy to edit from macros or external applications.
    /// </summary>
    public sealed class BirrDataCardValues
    {
        // These properties map directly to the Birr PDM data-card fields.
        // Defaults are kept here so objects automatically inherit the standard values unless
        // a macro or external system overrides them.
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

        /// <summary>
        /// Helper used when an object wants the default Birr values without overriding anything.
        /// </summary>
        public static BirrDataCardValues CreateDefault()
        {
            return new BirrDataCardValues();
        }

        /// <summary>
        /// Converts the object into the dictionary format required by the low-level PDM update code.
        ///
        /// If the drawing number is still empty, the generated PDM serial number can be used instead.
        /// </summary>
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

    /// <summary>
    /// Central service class for SolidWorks PDM operations.
    ///
    /// This class is responsible for:
    /// - logging into the vault
    /// - saving files into the vault
    /// - reading data-card values
    /// - writing data-card values
    /// </summary>
    public sealed class PdmModule
    {
        private IEdmVault5 _vault;
        private const string VaultRoot = @"C:\Users\kareem.salah\PDM\Birr Machines PDM";

        /// <summary>
        /// Opens the PDM vault session.
        ///
        /// Any workflow that wants to save to PDM must call this before SaveAsPdm().
        /// </summary>
        public void Login()
        {
            _vault = new EdmVault5();
            _vault.Login("kareem.salah", "976431852@KEmo", "Birr Machines PDM");

            if (!_vault.IsLoggedIn)
                throw new Exception("PDM Login failed.");

            Console.WriteLine("Logged into vault");
        }

        /// <summary>
        /// Copies an already existing local file into the vault and assigns a PDM serial number.
        /// </summary>
        public void AddExistingFileToPdm(string localFilePath, string subFolder = "60_Tests")
        {
            // Ask the vault for the next available serial number.
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

        /// <summary>
        /// Saves a live SolidWorks model directly into the PDM vault.
        ///
        /// This method also updates the vault data card when a BirrDataCardValues object is supplied.
        /// </summary>
        public string SaveAsPdm(ModelDoc2 swModel, string subFolder = "60_Tests", BirrDataCardValues dataCardValues = null)
        {
            // Generate the next serial number before saving.
            IEdmVault11 vault11 = (IEdmVault11)_vault;
            IEdmSerNoGen7 snGen = (IEdmSerNoGen7)vault11.CreateUtility(EdmUtility.EdmUtil_SerNoGen);
            snGen.GetSerialNumberNames(out string[] snNames);
            if (snNames?.Length == 0)
                throw new Exception("No Serial Numbers found.");

            var snValueObj = snGen.AllocSerNoValue(snNames[0], 0, string.Empty, 0, 0, 0, 0);
            string snString = snValueObj.Value;

            // Pick the correct SolidWorks extension from the document type.
            string extension = ".sldprt";
            int type = swModel.GetType();
            if (type == (int)swDocumentTypes_e.swDocASSEMBLY)
                extension = ".sldasm";
            else if (type == (int)swDocumentTypes_e.swDocDRAWING)
                extension = ".slddrw";

            // Build the final vault file path.
            string targetDir = Path.Combine(VaultRoot, subFolder);
            string fullPath = Path.Combine(targetDir, snString + extension);

            if (!Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            // Save from SolidWorks to the vault folder on disk first.
            swModel.SaveAs3(fullPath, 0, 1);

            // Then register that file inside the PDM database.
            IEdmFolder5 folder = _vault.GetFolderFromPath(targetDir);
            if (folder == null)
                throw new Exception("Vault folder object not found.");

            folder.AddFile(0, fullPath, string.Empty, 1);

            if (dataCardValues != null)
            {
                // Push the object's data-card values into the vault file.
                UpdateBirrDataCard(fullPath, dataCardValues.ToDictionary(snString));
            }

            Console.WriteLine($"Saved and Vaulted: {Path.GetFileName(fullPath)}");
            return fullPath;
        }

        /// <summary>
        /// Prints all non-empty data-card values for every configuration in a vault file.
        /// </summary>
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

        /// <summary>
        /// Writes a set of data-card values into a vault file.
        ///
        /// The method updates the "@" configuration and any real configurations found in the file.
        /// </summary>
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

            // Get a variable enumerator so we can write vault variables programmatically.
            IEdmEnumeratorVariable5 varEnum = (IEdmEnumeratorVariable5)file.GetEnumeratorVariable();
            IEdmVariableMgr5 varMgr = (IEdmVariableMgr5)_vault;

            // Always include the generic "@" configuration, then add any named configurations too.
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
                        // Vault variable names do not always match our code property names perfectly.
                        // So we search the vault variables and try to find the closest matching name.
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

            // Close and unlock so the values are committed back to the vault file.
            ((IEdmEnumeratorVariable8)varEnum).CloseFile(true);
            file.UnlockFile(0, "Automated data card sync");
        }

        /// <summary>
        /// Convenience helper that writes the built-in default Birr card values to a file.
        /// </summary>
        public void FillBirrDataCard(string relativePath)
        {
            BirrDataCardValues birrData = BirrDataCardValues.CreateDefault();
            UpdateBirrDataCard(relativePath, birrData.ToDictionary());
            Console.WriteLine($"Full Data Card populated for: {relativePath}");
        }
    }
}
