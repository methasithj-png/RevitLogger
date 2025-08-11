using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;

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

                // Compute and cache Revit version info
                RevitVersionInfo = $"{ctrl.VersionName} {ctrl.VersionNumber} {ctrl.SubVersionNumber}";

                // Create Ribbon UI
                try
                {
                    // Create or get tab
                    try { application.CreateRibbonTab(RibbonTabName); } catch { /* tab may already exist */ }
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
            catch
            {
                // ignore
            }
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
            catch
            {
                // ignore
            }
        }

        private static void OnDocumentClosing(object sender, DocumentClosingEventArgs e)
        {
            try
            {
                if (!SettingsManager.IsExportEnabled()) return;

                var doc = e.Document;
                if (doc == null) return;

                var now = DateTime.Now;

                // Build log data
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
                try { TaskDialog.Show("Project Close Logger", $"Logging failed: {ex.Message}"); } catch { /* ignore UI failures */ }
            }
            finally
            {
                // Cleanup per-document maps to avoid leaks
                try { if (DocumentOpenTimes.ContainsKey(e.Document)) DocumentOpenTimes.Remove(e.Document); } catch { }
                try { if (DocumentSyncCounts.ContainsKey(e.Document)) DocumentSyncCounts.Remove(e.Document); } catch { }
            }
        }

        private static string GetUserId(Application app)
        {
            try
            {
                // Revit username if set; fallback to OS username
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
            try
            {
                return collector.WhereElementIsNotElementType().GetElementCount();
            }
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

    internal sealed class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class
    {
        public bool Equals(T x, T y) => ReferenceEquals(x, y);
        public int GetHashCode(T obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }
}