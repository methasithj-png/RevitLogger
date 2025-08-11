using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;

namespace RevitProjectCloseLogger
{
    public class App : IExternalApplication
    {
        private static readonly Guid SessionId = Guid.NewGuid();
        private static readonly DateTime SessionStart = DateTime.Now;
        private static readonly Dictionary<Document, DateTime> DocumentOpenTimes = new Dictionary<Document, DateTime>(new ReferenceEqualityComparer<Document>());
        private static readonly Dictionary<Document, int> DocumentSyncCounts = new Dictionary<Document, int>(new ReferenceEqualityComparer<Document>());

        private static string RevitVersionInfo = string.Empty;
        private const string RibbonTabName = "Project Close Logger";
        private const string RibbonPanelName = "Logging";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                var ctrl = application.ControlledApplication;
                ctrl.DocumentOpened += OnDocumentOpened;
                ctrl.DocumentClosing += OnDocumentClosing;
                ctrl.DocumentSynchronizedWithCentral += OnDocumentSynchronizedWithCentral;

                RevitVersionInfo = $"{ctrl.VersionName} {ctrl.VersionNumber} {ctrl.SubVersionNumber}";

                try
                {
                    try { application.CreateRibbonTab(RibbonTabName); } catch { }
                    var panel = GetOrCreatePanel(application, RibbonTabName, RibbonPanelName);

                    var asmPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

                    var toggleBtnData = new PushButtonData(
                        name: "ToggleExport",
                        text: "Toggle Export",
                        assemblyName: asmPath,
                        className: typeof(ToggleExportCommand).FullName
                    )
                    {
                        ToolTip = "Enable/Disable export of project close logs to Excel-compatible CSV"
                    };

                    var btn = panel.AddItem(toggleBtnData) as PushButton;
                    if (btn != null)
                    {
                        btn.LongDescription = "Toggle automatic logging on project close. The setting persists per user.";
                    }
                }
                catch (Exception uiEx)
                {
                    TaskDialog.Show("Project Close Logger", $"Failed to create ribbon: {uiEx.Message}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Project Close Logger", $"Startup failed: {ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            var ctrl = application.ControlledApplication;
            ctrl.DocumentOpened -= OnDocumentOpened;
            ctrl.DocumentClosing -= OnDocumentClosing;
            ctrl.DocumentSynchronizedWithCentral -= OnDocumentSynchronizedWithCentral;
            DocumentOpenTimes.Clear();
            DocumentSyncCounts.Clear();
            return Result.Succeeded;
        }

        private static RibbonPanel GetOrCreatePanel(UIControlledApplication app, string tab, string panelName)
        {
            foreach (var panel in app.GetRibbonPanels(tab))
            {
                if (panel.Name.Equals(panelName, StringComparison.OrdinalIgnoreCase)) return panel;
            }
            return app.CreateRibbonPanel(tab, panelName);
        }

        private static void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            try
            {
                if (e.Document == null) return;
                if (!DocumentOpenTimes.ContainsKey(e.Document))
                {
                    DocumentOpenTimes[e.Document] = DateTime.Now;
                }
            }
            catch { }
        }

        private static void OnDocumentSynchronizedWithCentral(object sender, DocumentSynchronizedWithCentralEventArgs e)
        {
            try
            {
                var doc = e.Document;
                if (doc == null) return;
                if (!DocumentSyncCounts.ContainsKey(doc))
                {
                    DocumentSyncCounts[doc] = 0;
                }
                DocumentSyncCounts[doc]++;
            }
            catch { }
        }

        private static void OnDocumentClosing(object sender, DocumentClosingEventArgs e)
        {
            try
            {
                if (!SettingsManager.IsExportEnabled()) return;

                var doc = e.Document;
                if (doc == null) return;

                var now = DateTime.Now;

                var data = new LogData
                {
                    UserId = GetUserId(doc.Application),
                    SessionId = SessionId.ToString(),
                    Date = now,
                    ProjectName = GetProjectParameterOrEmpty(doc, BuiltInParameter.PROJECT_NAME),
                    ProjectNumber = GetProjectParameterOrEmpty(doc, BuiltInParameter.PROJECT_NUMBER),
                    ProjectFileName = GetProjectFileName(doc),
                    Action = "CloseProject",
                    LogStart = DocumentOpenTimes.ContainsKey(doc) ? DocumentOpenTimes[doc] : now,
                    LogEnd = now,
                    Duration = DocumentOpenTimes.ContainsKey(doc) ? (now - DocumentOpenTimes[doc]) : TimeSpan.Zero,
                    SessionDuration = now - SessionStart,
                    SynchronizedProjectsCount = DocumentSyncCounts.ContainsKey(doc) ? DocumentSyncCounts[doc] : 0,
                    UserComputer = Environment.MachineName,
                    RevitServicePackage = RevitVersionInfo,
                    Warnings = SafeGetWarningsCount(doc),
                    FileSize = SafeGetFileSize(doc),
                    Worksets = SafeGetWorksetCount(doc),
                    LinkedModels = SafeCount(doc, new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance))),
                    ImportedImages = SafeCount(doc, new FilteredElementCollector(doc).OfClass(typeof(ImageType))),
                    Views = SafeCountViews(doc),
                    Sheets = SafeCountSheets(doc),
                    ModelElements = SafeCountModelElements(doc),
                    ModelGroups = SafeCountByCategory(doc, BuiltInCategory.OST_IOSModelGroups),
                    DetailGroups = SafeCountByCategory(doc, BuiltInCategory.OST_IOSDetailGroups),
                    DesignOptions = SafeCount(doc, new FilteredElementCollector(doc).OfClass(typeof(DesignOption))),
                    Diroot = SafeGetDirectory(doc),
                    DesktopConnector = DetectDesktopConnector(doc),
                    ImportedCADFiles = SafeCount(doc, new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)))
                };

                ExcelExporter.AppendRow(data);
            }
            catch (Exception ex)
            {
                try { TaskDialog.Show("Project Close Logger", $"Logging failed: {ex.Message}"); } catch { }
            }
            finally
            {
                try { if (DocumentOpenTimes.ContainsKey(e.Document)) DocumentOpenTimes.Remove(e.Document); } catch { }
                try { if (DocumentSyncCounts.ContainsKey(e.Document)) DocumentSyncCounts.Remove(e.Document); } catch { }
            }
        }

        private static string GetUserId(Application app)
        {
            try
            {
                var revitUser = app.Username;
                if (!string.IsNullOrWhiteSpace(revitUser)) return revitUser;
            }
            catch { }
            return Environment.UserName;
        }

        private static string GetProjectParameterOrEmpty(Document doc, BuiltInParameter bip)
        {
            try
            {
                var pi = doc.ProjectInformation;
                if (pi == null) return string.Empty;
                var p = pi.get_Parameter(bip);
                return p != null ? p.AsString() ?? string.Empty : string.Empty;
            }
            catch { return string.Empty; }
        }

        private static string GetProjectFileName(Document doc)
        {
            try
            {
                var path = doc.PathName;
                if (string.IsNullOrWhiteSpace(path)) return doc.Title + ".rvt";
                return Path.GetFileName(path);
            }
            catch { return doc.Title; }
        }

        private static long SafeGetFileSize(Document doc)
        {
            try
            {
                var path = doc.PathName;
                if (string.IsNullOrWhiteSpace(path)) return 0;
                if (!File.Exists(path)) return 0;
                return new FileInfo(path).Length;
            }
            catch { return 0; }
        }

        private static int SafeGetWarningsCount(Document doc)
        {
            try
            {
                var warnings = doc.GetWarnings();
                return warnings != null ? warnings.Count : 0;
            }
            catch { return 0; }
        }

        private static int SafeGetWorksetCount(Document doc)
        {
            try
            {
                if (!doc.IsWorkshared) return 0;
                var col = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset);
                return col?.ToWorksets().Count() ?? 0;
            }
            catch { return 0; }
        }

        private static int SafeCount(Document doc, FilteredElementCollector collector)
        {
            try { return collector.WhereElementIsNotElementType().GetElementCount(); }
            catch { return 0; }
        }

        private static int SafeCountByCategory(Document doc, BuiltInCategory bic)
        {
            try
            {
                return new FilteredElementCollector(doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .GetElementCount();
            }
            catch { return 0; }
        }

        private static int SafeCountViews(Document doc)
        {
            try
            {
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Count(v => v != null && !v.IsTemplate);
            }
            catch { return 0; }
        }

        private static int SafeCountSheets(Document doc)
        {
            try
            {
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .WhereElementIsNotElementType()
                    .GetElementCount();
            }
            catch { return 0; }
        }

        private static int SafeCountModelElements(Document doc)
        {
            try
            {
                return new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(e => e.Category != null && e.Category.CategoryType == CategoryType.Model)
                    .GetElementCount();
            }
            catch { return 0; }
        }

        private static string SafeGetDirectory(Document doc)
        {
            try
            {
                var path = doc.PathName;
                if (string.IsNullOrWhiteSpace(path)) return string.Empty;
                return Path.GetDirectoryName(path) ?? string.Empty;
            }
            catch { return string.Empty; }
        }

        private static bool DetectDesktopConnector(Document doc)
        {
            try
            {
                var path = doc.PathName ?? string.Empty;
                path = path.ToLowerInvariant();
                return path.Contains("accdocs") || path.Contains("autodesk docs") || path.Contains("bim 360");
            }
            catch { return false; }
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class ToggleExportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                bool enabled = SettingsManager.IsExportEnabled();
                bool newValue = !enabled;
                SettingsManager.SetExportEnabled(newValue);

                var td = new TaskDialog("Project Close Logger")
                {
                    MainInstruction = newValue ? "Export ENABLED" : "Export DISABLED",
                    MainContent = "This setting controls whether a row is appended to the Excel-compatible CSV when a project is closed.",
                    CommonButtons = TaskDialogCommonButtons.Close
                };
                td.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    internal static class SettingsManager
    {
        private const string SettingsFileName = "settings.json";
        private const string AppFolderName = "RevitProjectCloseLogger";

        private class Settings
        {
            public bool ExportEnabled { get; set; } = true; // default ON
        }

        private static string GetAppFolder()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var path = Path.Combine(appData, AppFolderName);
            Directory.CreateDirectory(path);
            return path;
        }

        private static string GetSettingsPath()
        {
            return Path.Combine(GetAppFolder(), SettingsFileName);
        }

        public static bool IsExportEnabled()
        {
            try
            {
                var path = GetSettingsPath();
                if (!File.Exists(path)) return true; // default enabled
                var json = File.ReadAllText(path, Encoding.UTF8);
                var settings = JsonSerializer.Deserialize<Settings>(json);
                return settings?.ExportEnabled ?? true;
            }
            catch { return true; }
        }

        public static void SetExportEnabled(bool enabled)
        {
            try
            {
                var settings = new Settings { ExportEnabled = enabled };
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(GetSettingsPath(), json, Encoding.UTF8);
            }
            catch { }
        }

        public static string GetLogsFolder()
        {
            var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var folder = Path.Combine(docs, AppFolderName, "Logs");
            Directory.CreateDirectory(folder);
            return folder;
        }
    }

    internal class LogData
    {
        public string UserId { get; set; }
        public string SessionId { get; set; }
        public DateTime Date { get; set; }
        public string ProjectName { get; set; }
        public string ProjectNumber { get; set; }
        public string ProjectFileName { get; set; }
        public string Action { get; set; }
        public DateTime LogStart { get; set; }
        public DateTime LogEnd { get; set; }
        public TimeSpan Duration { get; set; }
        public TimeSpan SessionDuration { get; set; }
        public int SynchronizedProjectsCount { get; set; }
        public string UserComputer { get; set; }
        public string RevitServicePackage { get; set; }
        public int Warnings { get; set; }
        public long FileSize { get; set; }
        public int Worksets { get; set; }
        public int LinkedModels { get; set; }
        public int ImportedImages { get; set; }
        public int Views { get; set; }
        public int Sheets { get; set; }
        public int ModelElements { get; set; }
        public int ModelGroups { get; set; }
        public int DetailGroups { get; set; }
        public int DesignOptions { get; set; }
        public string Diroot { get; set; }
        public bool DesktopConnector { get; set; }
        public int ImportedCADFiles { get; set; }
    }

    internal static class ExcelExporter
    {
        private static readonly string[] Header = new[]
        {
            "UserId","SessionId","Date","ProjectName","ProjectNumber","ProjectFileName","Action","LogStart","LogEnd","DurationSeconds","SessionDurationSeconds","SynchronizedProjectsCount","UserComputer","RevitServicePackage","Warnings","FileSizeBytes","Worksets","LinkedModels","ImportedImages","Views","Sheets","ModelElements","ModelGroups","DetailGroups","DesignOptions","Diroot","DesktopConnector","ImportedCADFiles"
        };

        public static void AppendRow(LogData data)
        {
            var folder = SettingsManager.GetLogsFolder();
            var file = Path.Combine(folder, $"RevitLog_{DateTime.Now:yyyyMM}.csv");

            var sb = new StringBuilder();

            if (!File.Exists(file))
            {
                sb.AppendLine(string.Join(",", Header));
            }

            var row = new string[]
            {
                Csv(data.UserId),
                Csv(data.SessionId),
                Csv(data.Date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
                Csv(data.ProjectName),
                Csv(data.ProjectNumber),
                Csv(data.ProjectFileName),
                Csv(data.Action),
                Csv(data.LogStart.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
                Csv(data.LogEnd.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)),
                Csv(((long)data.Duration.TotalSeconds).ToString(CultureInfo.InvariantCulture)),
                Csv(((long)data.SessionDuration.TotalSeconds).ToString(CultureInfo.InvariantCulture)),
                Csv(data.SynchronizedProjectsCount.ToString(CultureInfo.InvariantCulture)),
                Csv(data.UserComputer),
                Csv(data.RevitServicePackage),
                Csv(data.Warnings.ToString(CultureInfo.InvariantCulture)),
                Csv(data.FileSize.ToString(CultureInfo.InvariantCulture)),
                Csv(data.Worksets.ToString(CultureInfo.InvariantCulture)),
                Csv(data.LinkedModels.ToString(CultureInfo.InvariantCulture)),
                Csv(data.ImportedImages.ToString(CultureInfo.InvariantCulture)),
                Csv(data.Views.ToString(CultureInfo.InvariantCulture)),
                Csv(data.Sheets.ToString(CultureInfo.InvariantCulture)),
                Csv(data.ModelElements.ToString(CultureInfo.InvariantCulture)),
                Csv(data.ModelGroups.ToString(CultureInfo.InvariantCulture)),
                Csv(data.DetailGroups.ToString(CultureInfo.InvariantCulture)),
                Csv(data.DesignOptions.ToString(CultureInfo.InvariantCulture)),
                Csv(data.Diroot),
                Csv(data.DesktopConnector ? "1" : "0"),
                Csv(data.ImportedCADFiles.ToString(CultureInfo.InvariantCulture))
            };

            sb.AppendLine(string.Join(",", row));

            File.AppendAllText(file, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        private static string Csv(string value)
        {
            if (value == null) return string.Empty;
            var needsQuotes = value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r");
            var escaped = value.Replace("\"", "\"\"");
            return needsQuotes ? $"\"{escaped}\"" : escaped;
        }
    }

    internal sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class
    {
        public bool Equals(T x, T y) => ReferenceEquals(x, y);
        public int GetHashCode(T obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }
}