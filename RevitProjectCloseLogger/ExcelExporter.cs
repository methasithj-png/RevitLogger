using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace RevitProjectCloseLogger
{
    internal static class ExcelExporter
    {
        private static readonly string[] Header = new[]
        {
            "UserId",
            "SessionId",
            "Date",
            "ProjectName",
            "ProjectNumber",
            "ProjectFileName",
            "Action",
            "LogStart",
            "LogEnd",
            "DurationSeconds",
            "SessionDurationSeconds",
            "SynchronizedProjectsCount",
            "UserComputer",
            "RevitServicePackage",
            "Warnings",
            "FileSizeBytes",
            "Worksets",
            "LinkedModels",
            "ImportedImages",
            "Views",
            "Sheets",
            "ModelElements",
            "ModelGroups",
            "DetailGroups",
            "DesignOptions",
            "Diroot",
            "DesktopConnector",
            "ImportedCADFiles"
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
}