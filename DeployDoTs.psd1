@{
    RootModule = "DeployDoTs.ps1"
    RequiredModules = @("ImportExcel", "PSTools")
    ModulesToImport = @("InvokeDoTs.ps1", "DeliveryOptimizationTroubleshooter.ps1")
    NestedModules = @("ImportExcel.psd1", "PSTools.psd1")
    Scripts = @("DeployDoTs.ps1", "InvokeDoTs.ps1", "DeliveryOptimizationTroubleshooter.ps1")
    FunctionsToExport = @("InvokeDoTs")
    RequiredAssemblies = @()
    TypesToProcess = @()
    FormatsToProcess = @()
    CmdletsToExport = @()
    AliasesToExport = @()
    VariablesToExport = @()
    DscResourcesToExport = @()
    ModuleVersion = "1.0.0"
    GUID = "12345678-1234-1234-1234-123456789012"
    Author = "Your Name"
    CompanyName = "Your Company"
    Copyright = "Your Company. All rights reserved."
    Description = "Comprehensive Delivery Optimization troubleshooting report in an Excel workbook."
    FileList = @(
        "Deploy-Do.ps1",
        "Invoke-DoTroubleshooter.ps1",
        "DeliveryOptimizationTroubleshooter.ps1",
        "Modules\ImportExcel\ImportExcel.psd1",
        "Modules\PSTools\PSTools.psd1"
    )
    RequiredScripts = @(
        "Modules\ImportExcel\ImportExcel.psm1",
        "Modules\PSTools\PSTools.psm1"
    )
}