@{
    # Module manifest properties
    ModuleVersion = "1.0.0"
    GUID = "9ae54ad3-e2c4-45f1-a4e0-319ce0b31a6d"
    Author = "ETSS_WinOps"
    CompanyName = "BAH"
    Description = "Delivery Optimization Troubleshooter for Windows - Generates comprehensive reports in Excel."
    
    # Module loading and dependency specifications
    RequiredModules = @("ImportExcel")
    # NestedModules = @("Modules\ImportExcel\7.8.10\ImportExcel.psd1")  # Commented out to avoid circular reference
    Scripts = @("DeployDoTs.ps1", "InvokeDoTs.ps1", "Scripts\DeliveryOptimizationTroubleshooter.ps1")
    
    # Export specifications
    FunctionsToExport = @("InvokeDoTs")
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    DscResourcesToExport = @()
    
    # Additional specifications
    RequiredAssemblies = @()
    TypesToProcess = @()
    FormatsToProcess = @()
    
    # File list for packaging - ensure all files are included
    FileList = @(
        "DeployDoTs.ps1",
        "InvokeDoTs.ps1",
        "PowerShell-7.4.0-win-x64.zip",
        "Scripts\DeliveryOptimizationTroubleshooter.ps1",
        "Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml",
        "Modules\ImportExcel\7.8.10\ImportExcel.psd1",
        "Modules\ImportExcel\7.8.10\ImportExcel.psm1",
        "Modules\ImportExcel\7.8.10\EPPlus.dll",
        "Modules\ImportExcel\7.8.10\en\ImportExcel-help.xml",
        "Modules\ImportExcel\7.8.10\en\Strings.psd1",
        "PSTools\PsExec64.exe",
        "PSTools\PsExec.exe",
        "PSTools\psfile.exe",
        "PSTools\psfile64.exe",
        "PSTools\PsGetsid.exe",
        "PSTools\PsGetsid64.exe",
        "PSTools\PsInfo.exe",
        "PSTools\PsInfo64.exe",
        "PSTools\pskill.exe",
        "PSTools\pskill64.exe",
        "PSTools\pslist.exe",
        "PSTools\pslist64.exe",
        "PSTools\PsLoggedon.exe",
        "PSTools\PsLoggedon64.exe",
        "PSTools\psloglist.exe",
        "PSTools\psloglist64.exe",
        "PSTools\pspasswd.exe",
        "PSTools\pspasswd64.exe",
        "PSTools\psping.exe",
        "PSTools\psping64.exe",
        "PSTools\PsService.exe",
        "PSTools\PsService64.exe",
        "PSTools\psshutdown.exe",
        "PSTools\psshutdown64.exe",
        "PSTools\pssuspend.exe",
        "PSTools\pssuspend64.exe",
        "Resources\app-icon.ico"
    )
    
    RequiredScripts = @(
        "Scripts\DeliveryOptimizationTroubleshooter.ps1",
        "Modules\ImportExcel\7.8.10\ImportExcel.psm1"
    )
    
    # PowerShell Pro Tools packaging configuration
    PackageConfiguration = @{
        Enabled = $true
        DotNetVersion = "net8.0"
        PowerShellVersion = "7.4.0"
        PackageType = "Console"
        HideConsoleWindow = $false  # Keep console visible for troubleshooting tool
        RequireElevation = $true    # Require admin rights
        ProductVersion = "1.0.0"
        FileVersion = "1.0.0.0"
        FileDescription = "Delivery Optimization Troubleshooter"
        ProductName = "DOTroubleshooterWin32"
        Copyright = "© 2025 BAH_ETSS"
        CompanyName = "BAH_ETSS_WinOps"
        Platform = "x64"
        RuntimeIdentifier = "win-x64"
        Host = "Default"
        HighDPISupport = $true
        IconFile = "Resources\app-icon.ico"
        OutputPath = ".\bin\Release"
        PublishSingleFile = $true
        SelfContained = $true
        
        # Entry point configuration - specify DeployDoTs.ps1 as the main script
        EntryPoint = "DeployDoTs.ps1"
        
        # Resource embedding configuration
        EmbeddedResources = @(
            @{
                Name = "InvokeDoTs.ps1"
                Path = "InvokeDoTs.ps1"
            },
            @{
                Name = "Scripts\DeliveryOptimizationTroubleshooter.ps1"
                Path = "Scripts\DeliveryOptimizationTroubleshooter.ps1"
            },
            @{
                Name = "Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml"
                Path = "Scripts\DeliveryOptimizationTroubleshooter_InstalledScriptInfo.xml"
            },
            @{
                Name = "PowerShell-7.4.0-win-x64.zip"
                Path = "PowerShell-7.4.0-win-x64.zip"
            },
            @{
                Name = "Resources\app-icon.ico"
                Path = "Resources\app-icon.ico"
            }
        )
        
        # Include PSTools directory as embedded resources
        IncludeDirectories = @(
            @{
                Name = "PSTools"
                Path = "PSTools"
            },
            @{
                Name = "Modules\ImportExcel\7.8.10"
                Path = "Modules\ImportExcel\7.8.10"
            }
        )
    }
}