@{
  # The main entry point of your package - the root script to package.
  Root       = "D:\do\DOTroubleshooterWin32"
  # The output directory where the packaged executable will be created.
  OutputPath = "D:\do\DOTroubleshooterWin32\out"
  # Package options determine how the executable is built.
  Package = @{
      Enabled             = $true                      # Enable packaging into an executable.
      DotNetVersion       = 'net8.0'                   # Target .NET version (for PowerShell 7.4.x).
      PowerShellVersion   = '7.4.0'                    # Target PowerShell version.
      PackageType         = 'Console'                  # Package type: "Console" or "Service".
      HideConsoleWindow   = $true                      # Hide the console window when running.
      RequireElevation    = $true                      # Require elevation to match script behavior.
      ProductVersion      = '1.0.0'                    # Product version.
      FileDescription     = 'DO Troubleshooter Launcher'
      ProductName         = 'DO Troubleshooter'
      Copyright           = '© 2025 BAH_ETSS'
      Platform            = 'x64'                      # Target architecture: x86 or x64.
      RuntimeIdentifier   = 'win-x64'                  # .NET runtime identifier.
      Host                = 'Default'                  # The host to use.
      HighDPISupport      = $true                      # Enable high DPI support.
      Lightweight         = $true                      # Use a lightweight executable.
      DisableQuickEdit    = $true                      # Disable quick edit mode in the console.
      Resources = [string[]]@(
            "D:\do\DOTroubleshooterWin32\Deploy-Do.ps1",
            "D:\do\DOTroubleshooterWin32\Invoke-DoTroubleshooter.ps1",
            "D:\do\DOTroubleshooterWin32\PowerShell-7.5.0-win-x64.zip",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\ImportExcel.psd1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\ImportExcel.psm1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\EPPlus.dll",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\en\ImportExcel-help.xml",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\en\Strings.psd1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Private\ArgumentCompletion.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Private\Invoke-ExcelReZipFile.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Add-ConditionalFormatting.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Add-ExcelChart.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Add-ExcelDataValidationRule.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Add-ExcelName.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Add-ExcelTable.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Add-PivotTable.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Add-Worksheet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Close-ExcelPackage.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Compare-Worksheet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Convert-ExcelRangeToImage.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelData.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelSheet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelToSQLInsert.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\ConvertTo-ExcelXlsx.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Copy-ExcelWorksheet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Enable-ExcelAutoFilter.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Enable-ExcelAutofit.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Expand-NumberFormat.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Export-Excel.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-ExcelColumnName.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-ExcelFileSchema.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-ExcelFileSummary.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-ExcelSheetDimensionAddress.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-ExcelSheetInfo.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-ExcelWorkbookInfo.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-HtmlTable.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-Range.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Get-XYRange.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Import-Excel.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Import-Html.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Import-UPS.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Import-USPS.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Invoke-ExcelQuery.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Invoke-Sum.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Join-Worksheet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Merge-MultipleSheets.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Merge-Worksheet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\New-ConditionalFormattingIconSet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\New-ConditionalText.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\New-ExcelChartDefinition.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\New-ExcelStyle.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\New-PivotTableDefinition.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\New-PSItem.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Open-ExcelPackage.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Read-Clipboard.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Read-OleDbData.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Remove-Worksheet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Select-Worksheet.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Send-SQLDataToExcel.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Set-CellComment.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Set-CellStyle.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Set-ExcelColumn.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Set-ExcelRange.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Set-ExcelRow.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Set-WorksheetProtection.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Public\Update-FirstObjectProperties.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Add-Subtotals.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Export-charts.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\GetExcelTable.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Plot.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Charting\Charting.ps1",
            "D:\do\DOTroubleshooterWin32\Modules\ImportExcel\7.8.10\Pivot\Pivot.ps1",
            "D:\do\DOTroubleshooterWin32\PSTools\PsExec.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsExec64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\psfile.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\psfile64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsGetsid.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsGetsid64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsInfo.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsInfo64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\pskill.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\pskill64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\pslist.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\pslist64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsLoggedon.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsLoggedon64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\psloglist.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\psloglist64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\pspasswd.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\pspasswd64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\psping.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\psping64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsService.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\PsService64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\psshutdown.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\psshutdown64.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\pssuspend.exe",
            "D:\do\DOTroubleshooterWin32\PSTools\pssuspend64.exe",
            "D:\do\DOTroubleshooterWin32\Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml",
            "D:\do\DOTroubleshooterWin32\Scripts\DeliveryOptimizationTroubleshooter.ps1"
      )
      OutputName          = 'DoTroubleshooter'         # The name of the output executable.
  }
  # Bundle options for including multiple scripts or modules into the package.
  Bundle = @{
      Enabled       = $true                          # Enable bundling.
      Modules       = $true                          # Bundle modules referenced by the script.
      NestedModules = $true                          # Include nested modules, if present.
    IgnoredModules = [string[]]@()                          # Modules to exclude from bundling.
  }
}
