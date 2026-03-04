# SwAutomationApp

Windows-only SOLIDWORKS automation app with SOLIDWORKS PDM integration.

## What Is Implemented Now
- Connect to SOLIDWORKS (`SldWorks`) and run it visible.
- Connect/login to SOLIDWORKS PDM vault.
- Generate CAD files from code:
  - Stator sheet (`Create_stator_sheet`)
  - Skeleton part (`CreateSkeleton`)
- Save generated files directly into PDM using serial-number-driven names.

## Current Runtime Flow
Entry point: `SwAutomationApp/Program.cs`

1. Start SOLIDWORKS.
2. Create `PdmModule`.
3. Create `Part` and inject `PdmModule` into it.
4. Login to PDM.
5. Generate and vault files:
   - Stator sheet
   - Skeleton

## Current Project Structure
```text
SwAutomationApp/
  Program.cs
  Part.cs
  Pdm.cs
  Assembly.cs
  SwAutomationApp.csproj
```

## PDM Integration (Latest)
File: `SwAutomationApp/Pdm.cs`

- `Login()`:
  - Uses `EdmVault5.Login(...)`
  - Verifies `_vault.IsLoggedIn`

- `SaveAsPdm(ModelDoc2 swModel, string subFolder = "60_Tests")`:
  - Self-contained flow (all required steps are inside this method).
  - Generates a new serial number from PDM (`IEdmSerNoGen7`).
  - Resolves extension from SOLIDWORKS document type (`.sldprt`, `.sldasm`, `.slddrw`).
  - Builds target path in local vault view.
  - Saves model via `SaveAs3`.
  - Registers file in vault folder using `AddFile`.
  - Returns the full vaulted path.

- `AddExistingFileToPdm(...)`:
  - Copies an existing local file to the target vault folder.
  - Assigns serial-number-based file name.
  - Registers it in PDM with `AddFile`.

## Part Class Updates
File: `SwAutomationApp/Part.cs`

- `Part` now receives `PdmModule` in constructor:
  - `Part(SldWorks swApp, PdmModule pdm)`
- `Create_stator_sheet(...)` and `CreateSkeleton(...)` call `_pdm.SaveAsPdm(...)`.
- Methods return the generated vaulted filename (not hardcoded names).

## Program.cs Updates
File: `SwAutomationApp/Program.cs`

- Uses namespace `SwAutomation.Pdm`.
- Creates and wires:
  - `var pdm = new SwAutomation.Pdm.PdmModule();`
  - `var myPart = new Part(swApp, pdm);`
- Calls `pdm.Login()` before saving CAD to vault.

## Requirements
- Windows machine.
- Installed SOLIDWORKS.
- Installed SOLIDWORKS PDM client + local vault view.
- .NET SDK compatible with project target (`net10.0-windows` in `SwAutomationApp.csproj`).
- COM interop references available at paths defined in `.csproj`:
  - `SolidWorks.Interop.sldworks.dll`
  - `SolidWorks.Interop.swconst.dll`
  - `EPDM.Interop.epdm.dll`

## Run
```powershell
cd "C:\Users\kareem.salah\Downloads\birr machines\birr machines\SwAutomationApp"
dotnet run
```

## Configuration Notes
- Vault root and credentials are currently hardcoded in `Pdm.cs`.
- Output subfolder is controlled in `Program.cs` (`outFolder`).
- Plane names in CAD operations currently use German naming (for example `Ebene vorne`).
