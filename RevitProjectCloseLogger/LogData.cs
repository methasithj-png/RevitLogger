using System;

namespace RevitProjectCloseLogger
{
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
}