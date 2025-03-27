================================================================================
Delivery Optimization (DO) Troubleshooter
================================================================================

📦 DESCRIPTION
This utility runs a comprehensive Delivery Optimization diagnostic on the system.
It outputs a detailed Excel report and CSV exports, including:

  • Service health and peering status
  • Endpoint reachability (Microsoft URLs)
  • Port testing (7680, 3544) for peer-to-peer connectivity
  • Microsoft Connected Cache (MCC) configuration
  • DNS-SD and DO Group ID validation
  • DO log analysis (Teams, Peer Sources, Failures)
  • WindowsUpdate.log integration with ETL conversion
  • Executive summary + prioritized recommendations

================================================================================
🚀 USAGE

Run the executable (`DOTUpdateCheck.exe`) or the launcher script (.\Run-DOTUpdateCheck.ps1) with:

    DOTUpdateCheck.exe [-Show] [-OutputPath <path>] [-DiagnosticsZip <zippath>]
    .\Run-DOTUpdateCheck.ps1 [-Show] [-OutputPath <path>] [-DiagnosticsZip <zippath>]

Examples:
    DOTUpdateCheck.exe
    DOTUpdateCheck.exe -Show
    DOTUpdateCheck.exe -OutputPath "C:\Reports"
    DOTUpdateCheck.exe -DiagnosticsZip "C:\Temp\DiagnosticsData.zip" -Show
    
    or
    
    .\Run-DOTUpdateCheck.ps1 -Show
    .\Run-DOTUpdateCheck.ps1 -OutputPath "C:\Reports"
    .\Run-DOTUpdateCheck.ps1 -DiagnosticsZip "C:\Temp\DiagnosticsData.zip" -Show

PARAMETERS:

    -Show
        Automatically opens the Excel report upon completion.

    -OutputPath <string>
        Sets the output folder for reports.
        Defaults to the user's Desktop.
        
    -DiagnosticsZip <string>
        Specifies a diagnostics ZIP file to extract and analyze.
        Supports Intune diagnostics ZIPs and Windows update logs.

================================================================================
🧾 OUTPUT FILES

After execution, you'll receive:

  • DO_Report_<timestamp>.xlsx     → Full Excel workbook with all diagnostics
  • DO_Report_CSV_<timestamp>\     → CSV exports of raw diagnostic buffers
  • DO_Report_Summary_<timestamp>.txt → Plain-text executive summary

================================================================================
📁 EMBEDDED MODULES

This tool includes a bundled version of the ImportExcel PowerShell module.
No need for Internet access or prerequisites.

================================================================================
🔒 PERMISSIONS

Some diagnostics require Administrator privileges.
The tool will auto-elevate if needed.

================================================================================
🆘 SUPPORT & DOCS

• Delivery Optimization Docs:
  [https://learn.microsoft.com/en-us/windows/deployment/optimization/](https://learn.microsoft.com/en-us/windows/deployment/optimization/)

• Troubleshooting Reference:
  [https://learn.microsoft.com/en-us/windows/deployment/optimization/waas-delivery-optimization-setup](https://learn.microsoft.com/en-us/windows/deployment/optimization/waas-delivery-optimization-setup)

================================================================================
📌 NOTES

• Compatible with Windows PowerShell 5.1 and PowerShell 7+
• PowerShell 7.4.0 is embedded and used by default
• All files are extracted to a temporary folder and cleaned up automatically
• Automatically converts ETL logs to WindowsUpdate.log for analysis

================================================================================
