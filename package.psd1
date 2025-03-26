@{
    # Root script entry and output directory (using $PSScriptRoot for dynamic resolution)
    Root       = "$PSScriptRoot\Deploy-Do.ps1"
    OutputPath = "$PSScriptRoot\out"

    Package = @{
        Enabled             = $true
        DotNetVersion       = "net8.0"
        PowerShellVersion   = "7.4.0"
        PackageType         = "Console"
        HideConsoleWindow   = $false
        RequireElevation    = $true
        ProductVersion      = "1.0.0"
        FileVersion         = "1.0.0.0"
        FileDescription     = "Delivery Optimization Troubleshooter"
        ProductName         = "DO Troubleshooter"
        Copyright           = "© 2025 BAH_ETSS"
        CompanyName         = "BAH_ETSS_WinOps"
        Platform            = "x64"
        RuntimeIdentifier   = "win-x64"
        Host                = "Default"
        HighDPISupport      = $true
        Lightweight         = $true
        DisableQuickEdit    = $true
        Icon                = "$PSScriptRoot\Resources\app-icon.ico"

        # Explicitly cast Resources as [string[]] to meet PowerShell Pro Tools expectations.
        Resources = [string[]]@(
            "$PSScriptRoot\Deploy-Do.ps1",
            "$PSScriptRoot\Invoke-DoTroubleshooter.ps1",
            "$PSScriptRoot\PowerShell-7.5.0-win-x64.zip",
            "$PSScriptRoot\PowerShell-7.4.0-win-x64.zip",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\ImportExcel.psd1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\ImportExcel.psm1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\EPPlus.dll",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\en\ImportExcel-help.xml",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\en\Strings.psd1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Private\ArgumentCompletion.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Private\Invoke-ExcelReZipFile.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Add-ConditionalFormatting.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Add-ExcelChart.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Add-ExcelDataValidationRule.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Add-ExcelName.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Add-ExcelTable.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Add-PivotTable.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Add-Worksheet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Close-ExcelPackage.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Compare-Worksheet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Convert-ExcelRangeToImage.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelData.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelSheet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\ConvertFrom-ExcelToSQLInsert.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\ConvertTo-ExcelXlsx.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Copy-ExcelWorksheet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Enable-ExcelAutoFilter.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Enable-ExcelAutofit.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Expand-NumberFormat.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Export-Excel.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-ExcelColumnName.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-ExcelFileSchema.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-ExcelFileSummary.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-ExcelSheetDimensionAddress.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-ExcelSheetInfo.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-ExcelWorkbookInfo.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-HtmlTable.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-Range.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Get-XYRange.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Import-Excel.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Import-Html.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Import-UPS.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Import-USPS.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Invoke-ExcelQuery.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Invoke-Sum.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Join-Worksheet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Merge-MultipleSheets.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Merge-Worksheet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\New-ConditionalFormattingIconSet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\New-ConditionalText.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\New-ExcelChartDefinition.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\New-ExcelStyle.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\New-PivotTableDefinition.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\New-PSItem.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Open-ExcelPackage.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Read-Clipboard.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Read-OleDbData.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Remove-Worksheet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Select-Worksheet.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Send-SQLDataToExcel.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Set-CellComment.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Set-CellStyle.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Set-ExcelColumn.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Set-ExcelRange.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Set-ExcelRow.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Set-WorksheetProtection.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Public\Update-FirstObjectProperties.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Add-Subtotals.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Export-charts.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\GetExcelTable.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Plot.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Charting\Charting.ps1",
            "$PSScriptRoot\Modules\ImportExcel\7.8.10\Pivot\Pivot.ps1",
            "$PSScriptRoot\PSTools\PsExec.exe",
            "$PSScriptRoot\PSTools\PsExec64.exe",
            "$PSScriptRoot\PSTools\psfile.exe",
            "$PSScriptRoot\PSTools\psfile64.exe",
            "$PSScriptRoot\PSTools\PsGetsid.exe",
            "$PSScriptRoot\PSTools\PsGetsid64.exe",
            "$PSScriptRoot\PSTools\PsInfo.exe",
            "$PSScriptRoot\PSTools\PsInfo64.exe",
            "$PSScriptRoot\PSTools\pskill.exe",
            "$PSScriptRoot\PSTools\pskill64.exe",
            "$PSScriptRoot\PSTools\pslist.exe",
            "$PSScriptRoot\PSTools\pslist64.exe",
            "$PSScriptRoot\PSTools\PsLoggedon.exe",
            "$PSScriptRoot\PSTools\PsLoggedon64.exe",
            "$PSScriptRoot\PSTools\psloglist.exe",
            "$PSScriptRoot\PSTools\psloglist64.exe",
            "$PSScriptRoot\PSTools\pspasswd.exe",
            "$PSScriptRoot\PSTools\pspasswd64.exe",
            "$PSScriptRoot\PSTools\psping.exe",
            "$PSScriptRoot\PSTools\psping64.exe",
            "$PSScriptRoot\PSTools\PsService.exe",
            "$PSScriptRoot\PSTools\PsService64.exe",
            "$PSScriptRoot\PSTools\psshutdown.exe",
            "$PSScriptRoot\PSTools\psshutdown64.exe",
            "$PSScriptRoot\PSTools\pssuspend.exe",
            "$PSScriptRoot\PSTools\pssuspend64.exe",
            "$PSScriptRoot\Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml",
            "$PSScriptRoot\Scripts\DeliveryOptimizationTroubleshooter.ps1"
        )

        OutputName = "DoTroubleshooter"
    }

    Bundle = @{
        Enabled        = $true
        Modules        = $true
        NestedModules  = $true
        IgnoredModules = @(
            ".git",
            ".gitignore",
            "PowerShellGet",
            "PackageManagement"
        )
    }

    Advanced = @{
        TrimWhitespace          = $true
        ObfuscateScript         = $false
        DecompilationProtection = $false
        DetectAzureModules      = $true
        CompressionLevel        = "Optimal"
        ErrorActionPreference   = "Stop"
        ProgressPreference      = "SilentlyContinue"
    }

    Signing = @{
        Enabled         = $true
        CertificatePath = "$PSScriptRoot\CodeSigning\CodeSigningCert.pfx"
        TimeStampServer = "http://timestamp.digicert.com"
    }
}
