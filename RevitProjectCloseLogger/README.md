### Revit Project Close Logger

Exports a row of project metrics when a Revit project is closed, into an Excel-compatible CSV. Includes a ribbon button to toggle export on/off.

#### Build
- Target .NET Framework: 4.8 (adjust for your Revit version as needed)
- References required:
  - RevitAPI.dll
  - RevitAPIUI.dll
- The `.csproj` references the default Revit 2024 installation path. Update paths if using another version.

#### Install
1. Build the project to produce `RevitProjectCloseLogger.dll`.
2. Copy the DLL to a folder, e.g. `%AppData%\Autodesk\Revit\Addins\2024\RevitProjectCloseLogger`.
3. Place `RevitProjectCloseLogger.addin` in `%AppData%\Autodesk\Revit\Addins\2024` and ensure its `Assembly` path points to the DLL.

#### Use
- Start Revit. A tab `Project Close Logger` with panel `Logging` and a button `Toggle Export` appears.
- Click `Toggle Export` to enable/disable logging.
- When closing a project, a row is appended to a monthly CSV: `Documents\RevitProjectCloseLogger\Logs\RevitLog_YYYYMM.csv`.

#### Columns
- UserId, SessionId, Date, ProjectName, ProjectNumber, ProjectFileName, Action, LogStart, LogEnd, DurationSeconds, SessionDurationSeconds, SynchronizedProjectsCount, UserComputer, RevitServicePackage, Warnings, FileSizeBytes, Worksets, LinkedModels, ImportedImages, Views, Sheets, ModelElements, ModelGroups, DetailGroups, DesignOptions, Diroot, DesktopConnector, ImportedCADFiles

Notes:
- CSV is Excel-compatible. If you require native `.xlsx`, integrate an OpenXML writer such as `DocumentFormat.OpenXml` and replace `ExcelExporter` accordingly.
- Some counts (images, model elements) are approximations based on available API classes and may vary with Revit versions.