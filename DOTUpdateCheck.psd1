@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'Run-DOTUpdateCheck.ps1'

    # Version number of this module.
    ModuleVersion = '1.0.0'

    # Supported PSEditions
    CompatiblePSEditions = @('Desktop', 'Core')

    # ID used to uniquely identify this module
    GUID = '05c27f9a-6a5e-44c3-8d37-2b2b3c8e9f7d'

    # Author of this module
    Author = 'DOTroubleshooter Team'

    # Company or vendor of this module
    CompanyName = 'BAH_ETSS_WinOps'

    # Copyright statement for this module
    Copyright = '(c) 2025 DOTroubleshooter. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'Delivery Optimization Troubleshooter tool that analyzes Windows Update issues by parsing diagnostic data, checking peer connectivity, and generating comprehensive Excel reports with actionable recommendations.'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1'

    # Functions to export from this module
    FunctionsToExport = @('Get-RootPath')

    # Cmdlets to export from this module
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module
    AliasesToExport = @()

    # List of all files packaged with this module
    FileList = @(
        'DOTUpdateCheck.ps1',
        'Run-DOTUpdateCheck.ps1',
        'pwsh\pwsh.exe',
        'Assets\dots.ico',
        'Modules\ImportExcel'
    )

    # Private data to pass to the module specified in RootModule
    PrivateData = @{
        PSData = @{
            Tags = @('Delivery', 'Optimization', 'Windows', 'Updates', 'Diagnostics', 'Intune', 'WUfB', 'Offline', 'Portable')
            ReleaseNotes = 'Initial release of Delivery Optimization Troubleshooter with portable PowerShell 7.4.0, Excel reporting capabilities, and peer connectivity diagnostics.'
            RequireLicenseAcceptance = $false
        }
        
        # PowerShellProTools specific packaging information
        PowerShellProTools = @{
            PackageType = 'Console'
            IconUri = 'Assets\dots.ico'
            MainScript = 'Run-DOTUpdateCheck.ps1'
            Recurse = $true  # Include all subfolders
            CompressionLevel = 'High'
            ExeName = 'DOTUpdateCheck.exe'
            RequireElevation = $false
            ExcludeModules = @()
            PowerShellVersion = '7.4'
            TargetFramework = 'net7.0-windows'
        }
    }
}
