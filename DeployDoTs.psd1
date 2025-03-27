@{
    RootModule = "DeployDoTs.ps1"
    RequiredModules = @("ImportExcel")
    ModulesToImport = @("InvokeDoTs.ps1")
    NestedModules = @("Modules\ImportExcel\7.8.10\ImportExcel.psd1")
    Scripts = @("DeployDoTs.ps1", "InvokeDoTs.ps1", "Scripts\DeliveryOptimizationTroubleshooter.ps1")
    FunctionsToExport = @("InvokeDoTs")
    RequiredAssemblies = @()
    TypesToProcess = @()
    FormatsToProcess = @()
    CmdletsToExport = @()
    AliasesToExport = @()
    VariablesToExport = @()
    DscResourcesToExport = @()
    ModuleVersion = "1.0.0"
    GUID = "9ae54ad3-e2c4-45f1-a4e0-319ce0b31a6d"
    Author = "ETSS_WinOps"
    CompanyName = "BAH"
    Description = "Delivery Optimization Troubleshooter for Windows - Generates comprehensive reports in Excel."
    FileList = @(
        "DeployDoTs.ps1",
        "InvokeDoTs.ps1",
        "PowerShell-7.4.0-win-x64.zip",
        "Scripts\DeliveryOptimizationTroubleshooter.ps1",
        "Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml",
        "Modules\ImportExcel\7.8.10\ImportExcel.psd1",
        "Modules\ImportExcel\7.8.10\ImportExcel.psm1",
        "PSTools\PsExec64.exe"
    )
    RequiredScripts = @(
        "Scripts\DeliveryOptimizationTroubleshooter.ps1",
        "Modules\ImportExcel\7.8.10\ImportExcel.psm1"
    )
    
    # Packaging configuration
    PackageConfiguration = @{
        Enabled = $true
        DotNetVersion = "net8.0"
        PowerShellVersion = "7.4.0"
        PackageType = "Console"
        HideConsoleWindow = $true
        RequireElevation = $false
        ProductVersion = "1.0.0"
        FileVersion = "1.0.0.0"
        FileDescription = "Delivery Optimization Troubleshooter"
        ProductName = "InvokeDoTs"
        Copyright = "© 2025 BAH_ETSS"
        CompanyName = "BAH_ETSS_WinOps"
        Platform = "x64"
        RuntimeIdentifier = "win-x64"
        Host = "Default"
        HighDPISupport = $true
        Lightweight = $true
        DisableQuickEdit = $true
        Icon = "$ScriptRoot\Resources\app-icon.ico"
    }
}
