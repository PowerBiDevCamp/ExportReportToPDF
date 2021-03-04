using ExportReportToPDF.Models;

namespace ExportReportToPDF {

  class Program {
    static void Main() {
      
      var workspace = PowerBiManager.GetAppWorkspace("Report Export Demo");
      var report = PowerBiManager.GetReport(workspace.Id, "Export Demo");

      PowerBiManager.ExportReport(workspace.Id, report.Id);
    }
  }
}
