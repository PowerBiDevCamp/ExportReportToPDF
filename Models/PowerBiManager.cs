using System;
using System.Collections.Generic;
using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using System.IO;

namespace ExportReportToPDF.Models {

  class PowerBiManager {

    private readonly static string[] requiredScopes = PowerBiPermissionScopes.ManageWorkspaceAssets;

    public static void GetAppWorkspaces() {
      PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      var workspaces = pbiClient.Groups.GetGroups().Value;
      foreach (var workspace in workspaces) {
        Console.WriteLine(workspace.Name);
      }
    }

    public static void GetAppWorkspacesAsAdmin() {
      PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      var workspaces = pbiClient.Groups.GetGroupsAsAdmin(top: 100).Value;
      foreach (var workspace in workspaces) {
        Console.WriteLine(workspace.Name);
      }
    }

    public static Group GetAppWorkspace(string WorkspaceName) {
      PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      var workspaces = pbiClient.Groups.GetGroups().Value;
      foreach (var workspace in workspaces) {
        if (workspace.Name.Equals(WorkspaceName))
          return workspace;
      }
      return null;
    }

    public static Dataset GetDataset(Guid WorkspaceId, string DatasetName) {
      PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);
      var datasets = pbiClient.Datasets.GetDatasetsInGroup(WorkspaceId).Value;
      foreach (var dataset in datasets) {
        if (dataset.Name.Equals(DatasetName)) {
          return dataset;
        }
      }
      return null;
    }

    public static Report GetReport(Guid WorkspaceId, string ReportName) {
      PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);
      var reports = pbiClient.Reports.GetReportsInGroup(WorkspaceId).Value;
      foreach (var report in reports) {
        if (report.Name.Equals(ReportName)) {
          return report;
        }
      }
      return null;
    }

    public static void ExportReport(Guid WorkspaceId, Guid ReportId) {

      PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      var exportRequest = new ExportReportRequest {
        Format = FileFormat.PDF,
        PowerBIReportConfiguration = new PowerBIReportExportConfiguration {
          ReportLevelFilters = new List<ExportFilter>() {
           new ExportFilter(filter: "Customers/State eq 'TX'")
          }
        }
      };

      Export export = pbiClient.Reports.ExportToFileInGroup(WorkspaceId, ReportId, exportRequest);

      string exportId = export.Id;

      do {
        System.Threading.Thread.Sleep(5000);
        export = pbiClient.Reports.GetExportToFileStatus(ReportId, exportId);
        Console.WriteLine("Getting export status - " + export.PercentComplete.ToString() + "% complete");
      } while (export.Status != ExportState.Succeeded && export.Status != ExportState.Failed);

      if (export.Status == ExportState.Failed) {
        Console.WriteLine("Export failed!");
      }

      if (export.Status == ExportState.Succeeded) {
        string fileName = @"c:\DevCamp\Report1.pdf";
        Console.WriteLine("Saving exported file to " + fileName);
        var exportStream = pbiClient.Reports.GetFileOfExportToFile(WorkspaceId, ReportId, exportId);
        FileStream fileStream = File.Create(fileName);
        exportStream.CopyTo(fileStream);
        fileStream.Close();
      }

    }

  public static void ExportVisual(Guid WorkspaceId, Guid ReportId, string PageName, string VisualName) {

      PowerBIClient pbiClient = TokenManager.GetPowerBiClient(requiredScopes);

      var exportRequest = new ExportReportRequest {        
        Format = FileFormat.PDF,        
        PowerBIReportConfiguration = new PowerBIReportExportConfiguration {
          Pages = new List<ExportReportPage>() {
            new ExportReportPage{ 
              PageName=PageName, 
              VisualName=VisualName
            }
          }
        }      
      };

      Export export = pbiClient.Reports.ExportToFileInGroup(WorkspaceId, ReportId, exportRequest);

      string exportId = export.Id;

      do {
        System.Threading.Thread.Sleep(5000);
        export = pbiClient.Reports.GetExportToFileStatus(ReportId, exportId);
        Console.WriteLine("Getting export status - " + export.PercentComplete.ToString() + "% complete");
      } while (export.Status != ExportState.Succeeded && export.Status != ExportState.Failed);

      if (export.Status == ExportState.Failed) {
        Console.WriteLine("Export failed!");
      }

      if (export.Status == ExportState.Succeeded) {
        string fileName = @"c:\DevCamp\Visual1.pdf";
        Console.WriteLine("Saving exported file to " + fileName);
        var exportStream = pbiClient.Reports.GetFileOfExportToFile(WorkspaceId, ReportId, exportId);
        FileStream fileStream = File.Create(fileName);
        exportStream.CopyTo(fileStream);
        fileStream.Close();
      }

    }

  }
}
