# DoTroubleshooter package.psd1

# Resolve root
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Build configuration
$global:BuildConfig = @{
    Root       = "$ScriptRoot\Deploy-Do.ps1"
    OutputPath = "$ScriptRoot\out"

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
        Icon                = "$ScriptRoot\Resources\app-icon.ico"

        Resources = @(
            "$ScriptRoot\Deploy-Do.ps1",
            "$ScriptRoot\Invoke-DoTroubleshooter.ps1",
            "$ScriptRoot\PowerShell-7.5.0-win-x64.zip",
            "$ScriptRoot\PowerShell-7.4.0-win-x64.zip",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\ImportExcel.psd1",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\ImportExcel.psm1",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\EPPlus.dll",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\en\ImportExcel-help.xml",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\en\Strings.psd1",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\Private\*.ps1",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\Public\*.ps1",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\*.ps1",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\Charting\Charting.ps1",
            "$ScriptRoot\Modules\ImportExcel\7.8.10\Pivot\Pivot.ps1",
            "$ScriptRoot\PSTools\*.exe",
            "$ScriptRoot\Scripts\DeliveryOptimizationTroubleshooter.ps1",
            "$ScriptRoot\Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml"
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
        CertificatePath = "$ScriptRoot\CodeSigning\CodeSigningCert.pfx"
        TimeStampServer = "http://timestamp.digicert.com"
    }
}