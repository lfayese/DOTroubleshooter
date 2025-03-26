# DoTroubleshooter package.psd1

# Use $PSScriptRoot for base directory
$baseDir = $PSScriptRoot

# Build configuration
$BuildConfig = @{
    Root       = Join-Path $baseDir "Deploy-Do.ps1"
    OutputPath = Join-Path $baseDir "out"

    Package = @{
        Enabled           = $true
        DotNetVersion     = "net8.0"
        PowerShellVersion = "7.4.0"
        PackageType       = "Console"
        HideConsoleWindow = $false
        RequireElevation  = $true
        ProductVersion    = "1.0.0"
        FileVersion      = "1.0.0.0"
        FileDescription  = "Delivery Optimization Troubleshooter"
        ProductName      = "DO Troubleshooter"
        Copyright        = "© 2025 BAH_ETSS"
        CompanyName      = "BAH_ETSS_WinOps"
        Platform         = "x64"
        RuntimeIdentifier = "win-x64"
        Host             = "Default"
        HighDPISupport   = $true
        Lightweight      = $true
        DisableQuickEdit = $true
        Icon             = Join-Path $baseDir "Resources\app-icon.ico"
        OutputName       = "DoTroubleshooter"
        Resources        = @(
            (Join-Path $baseDir "Deploy-Do.ps1"),
            (Join-Path $baseDir "Invoke-DoTroubleshooter.ps1"),
            (Join-Path $baseDir "PowerShell-7.4.0-win-x64.zip"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\ImportExcel.psd1"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\ImportExcel.psm1"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\EPPlus.dll"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\en\ImportExcel-help.xml"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\en\Strings.psd1"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\Private\*.ps1"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\Public\*.ps1"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\*.ps1"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\Charting\Charting.ps1"),
            (Join-Path $baseDir "Modules\ImportExcel\7.8.10\Pivot\Pivot.ps1"),
            (Join-Path $baseDir "PSTools\*.exe"),
            (Join-Path $baseDir "Scripts\DeliveryOptimizationTroubleshooter.ps1"),
            (Join-Path $baseDir "Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml")
        )
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
        Enabled         = $false
        CertificatePath = Join-Path $baseDir "CodeSigning\CodeSigningCert.pfx"
        TimeStampServer = "https://timestamp.digicert.com"
    }
}

# Return the build configuration
$BuildConfig
