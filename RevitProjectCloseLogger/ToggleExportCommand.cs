using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitProjectCloseLogger
{
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
}