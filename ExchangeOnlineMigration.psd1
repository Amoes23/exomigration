# Module manifest for ExchangeOnlineMigration module
@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'ExchangeOnlineMigration.psm1'
    
    # Version number of this module.
    ModuleVersion = '3.1.0'
    
    # Supported PSEditions
    CompatiblePSEditions = @('Desktop', 'Core')
    
    # ID used to uniquely identify this module
    GUID = '43b21913-6f1a-4ddd-aeb6-2d7a8b98a5a0'
    
    # Author of this module
    Author = 'Exchange Migration Team'
    
    # Company or vendor of this module
    CompanyName = 'IT&Care'
    
    # Copyright statement for this module
    Copyright = '(c) 2025. All rights reserved.'
    
    # Description of the functionality provided by this module
    Description = 'A comprehensive PowerShell module for migrating mailboxes between Exchange environments'
    
    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1'
    
    # Name of the PowerShell host required by this module
    # PowerShellHostName = ''
    
    # Minimum version of the PowerShell host required by this module
    # PowerShellHostVersion = ''
    
    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    DotNetFrameworkVersion = '4.7.2'
    
    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # ClrVersion = ''
    
    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''
    
    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        @{ModuleName = 'ExchangeOnlineManagement'; ModuleVersion = '3.0.0'},
        @{ModuleName = 'Microsoft.Graph'; ModuleVersion = '1.20.0'},
        @{ModuleName = 'Microsoft.Graph.Users'; ModuleVersion = '1.20.0'}
    )
    
    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()
    
    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @()
    
    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()
    
    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()
    
    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()
    
    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport = @(
        'Start-EXOMigration',
        'Test-EXOMailboxReadiness',
        'Invoke-EXOParallelValidation',
        'New-EXOMigrationReport',
        'Get-EXOMigrationStatus',
        'New-EXOMigrationBatch'
    )
    
    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport = @()
    
    # Variables to export from this module
    VariablesToExport = 'EXOMigrationErrorCodes'
    
    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    AliasesToExport = @(
        'Start-ExchangeMigration',
        'Test-MailboxReadiness'
    )
    
    # DSC resources to export from this module
    # DscResourcesToExport = @()
    
    # List of all modules packaged with this module
    # ModuleList = @()
    
    # List of all files packaged with this module
    # FileList = @()
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{
        PSData = @{
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @('Exchange', 'Migration', 'Office365', 'ExchangeOnline', 'Mailbox', 'Hybrid')
            
            # A URL to the license for this module.
            # LicenseUri = ''
            
            # A URL to the main website for this project.
            # ProjectUri = ''
            
            # A URL to an icon representing this module.
            # IconUri = ''
            
            # ReleaseNotes of this module
            ReleaseNotes = @'
# Version 3.1.0
- Added unified architecture for bidirectional migrations
- Consolidated duplicate functions for on-premises and cloud environments
- Added support for cross-tenant migrations
- Enhanced error handling and connection management
- Improved performance for large migrations
- Added comprehensive unified validation test catalog
- Enhanced reporting with directional migration information

# Version 3.0.0
- Complete rewrite as proper PowerShell module
- Added tiered validation levels (Basic, Standard, Comprehensive)
- Improved memory management for large migrations
- Added checkpointing and recovery mechanisms
- Modern authentication with certificate support
- External HTML template support
- Performance optimizations for large environments
'@
            
            # Prerelease string of this module
            # Prerelease = ''
            
            # Flag to indicate whether the module requires explicit user acceptance for install/update/save
            # RequireLicenseAcceptance = $false
            
            # External dependent modules of this module
            # ExternalModuleDependencies = @()
        } # End of PSData hashtable
    } # End of PrivateData hashtable
    
    # HelpInfo URI of this module
    # HelpInfoURI = ''
    
    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''
}
