@{
    # The main entry point of your package - the root script to package
    Root = ".\Deploy-Do.ps1"  # Changed to relative path for better portability
  
    # The output directory where the packaged executable will be created
    OutputPath = ".\out"      # Changed to relative path for better portability
  
    # Package options determine how the executable is built
    Package = @{
        Enabled             = $true
        DotNetVersion       = 'net8.0'                   # Latest .NET version for best performance
        PowerShellVersion   = '7.4.0'                    # Target PowerShell version
        PackageType         = 'Console'                  # Console application
        HideConsoleWindow   = $false                     # Changed to false for better troubleshooting visibility
        RequireElevation    = $true                      # Require admin rights
        ProductVersion      = '1.0.0'
        FileVersion         = '1.0.0.0'                  # Added explicit file version
        FileDescription     = 'Delivery Optimization Troubleshooter'  # More descriptive
        ProductName         = 'DO Troubleshooter'
        Copyright           = '© 2025 BAH_ETSS'
        CompanyName         = 'BAH_ETSS_WinOps'                 # Added company name
        Platform            = 'x64'                      # 64-bit only
        RuntimeIdentifier   = 'win-x64'                  # Windows 64-bit
        Host                = 'Default'
        HighDPISupport      = $true                      # Support high DPI displays
        Lightweight         = $true                      # Optimize for size
        DisableQuickEdit    = $true                      # Prevent accidental pausing
        Icon                = '.\Resources\app-icon.ico' # Optional: Add an icon if available     
  
        # Resources to include in the package - using relative paths for portability
        Resources = @(
            ".\Deploy-Do.ps1"
            ".\Invoke-DoTroubleshooter.ps1"
            ".\PowerShell-7.5.0-win-x64.zip"
            ".\PowerShell-7.4.0-win-x64.zip"
            ".\Modules\ImportExcel\7.8.10\ImportExcel.psd1"
            ".\Modules\ImportExcel\7.8.10\ImportExcel.psm1"
            ".\Modules\ImportExcel\7.8.10\EPPlus.dll"
            ".\Modules\ImportExcel\7.8.10\en\ImportExcel-help.xml"
            ".\Modules\ImportExcel\7.8.10\en\Strings.psd1"
            ".\Modules\ImportExcel\7.8.10\Private\ArgumentCompletion.ps1"
            ".\Modules\ImportExcel\7.8.10\Private\Invoke-ExcelReZipFile.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Add-ConditionalFormatting.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Add-ExcelChart.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Add-ExcelDataValidationRule.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Add-ExcelName.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Add-ExcelTable.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Add-PivotTable.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Add-Worksheet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Close-ExcelPackage.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Compare-Worksheet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Convert-ExcelRangeToImage.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelData.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelSheet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelToSQLInsert.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\ConvertTo-ExcelXlsx.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Copy-ExcelWorksheet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Enable-ExcelAutoFilter.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Enable-ExcelAutofit.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Expand-NumberFormat.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Export-Excel.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-ExcelColumnName.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-ExcelFileSchema.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-ExcelFileSummary.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-ExcelSheetDimensionAddress.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-ExcelSheetInfo.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-ExcelWorkbookInfo.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-HtmlTable.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-Range.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Get-XYRange.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Import-Excel.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Import-Html.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Import-UPS.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Import-USPS.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Invoke-ExcelQuery.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Invoke-Sum.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Join-Worksheet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Merge-MultipleSheets.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Merge-Worksheet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\New-ConditionalFormattingIconSet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\New-ConditionalText.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\New-ExcelChartDefinition.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\New-ExcelStyle.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\New-PivotTableDefinition.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\New-PSItem.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Open-ExcelPackage.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Read-Clipboard.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Read-OleDbData.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Remove-Worksheet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Select-Worksheet.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Send-SQLDataToExcel.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Set-CellComment.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Set-CellStyle.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Set-ExcelColumn.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Set-ExcelRange.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Set-ExcelRow.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Set-WorksheetProtection.ps1"
            ".\Modules\ImportExcel\7.8.10\Public\Update-FirstObjectProperties.ps1"
            ".\Modules\ImportExcel\7.8.10\Add-Subtotals.ps1"
            ".\Modules\ImportExcel\7.8.10\Export-charts.ps1"
            ".\Modules\ImportExcel\7.8.10\GetExcelTable.ps1"
            ".\Modules\ImportExcel\7.8.10\Plot.ps1"
            ".\Modules\ImportExcel\7.8.10\Charting\Charting.ps1"
            ".\Modules\ImportExcel\7.8.10\Pivot\Pivot.ps1"
            ".\PSTools\PsExec.exe"
            ".\PSTools\PsExec64.exe"
            ".\PSTools\psfile.exe"
            ".\PSTools\psfile64.exe"
            ".\PSTools\PsGetsid.exe"
            ".\PSTools\PsGetsid64.exe"
            ".\PSTools\PsInfo.exe"
            ".\PSTools\PsInfo64.exe"
            ".\PSTools\pskill.exe"
            ".\PSTools\pskill64.exe"
            ".\PSTools\pslist.exe"
            ".\PSTools\pslist64.exe"
            ".\PSTools\PsLoggedon.exe"
            ".\PSTools\PsLoggedon64.exe"
            ".\PSTools\psloglist.exe"
            ".\PSTools\psloglist64.exe"
            ".\PSTools\pspasswd.exe"
            ".\PSTools\pspasswd64.exe"
            ".\PSTools\psping.exe"
            ".\PSTools\psping64.exe"
            ".\PSTools\PsService.exe"
            ".\PSTools\PsService64.exe"
            ".\PSTools\psshutdown.exe"
            ".\PSTools\psshutdown64.exe"
            ".\PSTools\pssuspend.exe"
            ".\PSTools\pssuspend64.exe"
            ".\Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml"
            ".\Scripts\DeliveryOptimizationTroubleshooter.ps1"
        )
        OutputName = 'DoTroubleshooter'
    }
  
    # Bundle options for including multiple scripts or modules into the package
    Bundle = @{
        Enabled       = $true
        Modules       = $true
        NestedModules = $true
        IgnoredModules = @(
            '.git'
            '.gitignore'
            'PowerShellGet'
            'PackageManagement'  # Added common modules to ignore
        )
    }
  
    # Advanced options for fine-tuning the build process
    Advanced = @{
        TrimWhitespace        = $true                # Remove unnecessary whitespace
        ObfuscateScript       = $false               # Don't obfuscate for easier troubleshooting 
        DecompilationProtection = $false             # Don't add decompilation protection
        DetectAzureModules    = $true                # Detect and handle Azure modules properly   
        CompressionLevel      = 'Optimal'            # Balance between size and performance       
        ErrorActionPreference = 'Stop'               # Stop on errors during packaging
        ProgressPreference    = 'SilentlyContinue'   # Don't show progress bars during packaging  
    }
  
     # Signing options - these will be populated dynamically by the build script
     Signing = @{
      Enabled     = $true
      CertificatePath = ".\CodeSigning\CodeSigningCert.pfx"
      # Password will be provided by the build script
      TimeStampServer = "http://timestamp.digicert.com"
    }
  }
  