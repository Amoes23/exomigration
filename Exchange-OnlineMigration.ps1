<#
.SYNOPSIS
    Exchange Online Migration Script - Comprehensive migration tool for moving mailboxes from on-premises Exchange to Exchange Online.

.DESCRIPTION
    This script automates the process of migrating mailboxes from on-premises Exchange to Exchange Online.
    It performs extensive pre-migration checks, creates migration batches, logs all activities,
    and generates an HTML report showing the migration status and results.
    
    The script includes:
    - Dependency validation for required PowerShell modules
    - Configuration from external JSON file
    - Comprehensive pre-migration validation
    - Parallel processing of mailbox validation for improved performance
    - Dry-run capability to test without making changes
    - Detailed HTML reporting with actionable recommendations
    - Secure handling of sensitive information

.PARAMETER BatchFilePath
    Path to the CSV file containing mailboxes to migrate. Must have "EmailAddress" as a header.

.PARAMETER ConfigPath
    Path to the JSON configuration file. If not specified, defaults to .\Config\ExchangeMigrationConfig.json

.PARAMETER TargetDeliveryDomain
    The mail.onmicrosoft.com domain for your tenant. If specified, overrides the value in the config file.

.PARAMETER MigrationEndpointName
    Name of the migration endpoint to use for the migration. If specified, overrides the value in the config file.

.PARAMETER LogPath
    Path where logs will be saved. Default is .\Logs\

.PARAMETER ReportPath
    Path where the HTML report will be saved. Default is .\Reports\

.PARAMETER NotificationEmails
    Email addresses to notify about migration status. If specified, overrides the value in the config file.

.PARAMETER MaxConcurrentMailboxes
    Maximum number of mailboxes to process in parallel. Default is 5.

.PARAMETER DryRun
    If specified, validates everything but doesn't create the migration batch.

.PARAMETER Force
    If specified, attempts to create the migration batch even if validation errors are present.

.EXAMPLE
    .\Exchange-OnlineMigration.ps1 -BatchFilePath ".\Batches\Batch01.csv" -ConfigPath ".\Config\MigrationConfig.json"

.EXAMPLE
    .\Exchange-OnlineMigration.ps1 -BatchFilePath ".\Batches\Batch01.csv" -DryRun

.EXAMPLE
    .\Exchange-OnlineMigration.ps1 -BatchFilePath ".\Batches\Batch01.csv" -TargetDeliveryDomain "contoso.mail.onmicrosoft.com" -MigrationEndpointName "Hybrid Migration Endpoint" -MaxConcurrentMailboxes 10

.NOTES
    Author: Amadeus van Daalen
    Version: 2.0
    Date: March 9, 2025
    Requires: Exchange Online PowerShell Module, Microsoft Graph PowerShell SDK
#>

#Requires -Version 5.1

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$BatchFilePath,
    
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = ".\Config\ExchangeMigrationConfig.json",
    
    [Parameter(Mandatory = $false)]
    [string]$TargetDeliveryDomain,
    
    [Parameter(Mandatory = $false)]
    [string]$MigrationEndpointName,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath,
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath,
    
    [Parameter(Mandatory = $false)]
    [string[]]$NotificationEmails,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 20)]
    [int]$MaxConcurrentMailboxes = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

#region Script Variables and Configuration
# Script metadata
$script:ScriptVersion = "2.0"
$script:ScriptName = "Exchange-OnlineMigration"

$script:ErrorCodes = @{
    # Connection and initialization errors
    "ERR001" = "Failed to load required PowerShell modules"
    "ERR002" = "Failed to load configuration from file"
    "ERR003" = "Failed to initialize environment"
    "ERR004" = "Failed to connect to Exchange Online"
    "ERR005" = "Failed to connect to Microsoft Graph"
    "ERR006" = "Failed to validate migration endpoint"
    
    # Batch and file errors
    "ERR007" = "Failed to import migration batch CSV"
    "ERR008" = "Batch CSV doesn't contain required 'EmailAddress' column"
    "ERR009" = "Migration batch with same name already exists"
    "ERR010" = "Failed to create migration batch"
    
    # Mailbox configuration errors
    "ERR011" = "Invalid mailbox format"
    "ERR012" = "Missing required license"
    "ERR013" = "License provisioning error"
    "ERR014" = "Missing onmicrosoft.com email address"
    "ERR015" = "Domain not verified in tenant"
    "ERR016" = "Pending move request detected"
    "ERR017" = "No mailboxes ready for migration"
    "ERR018" = "Failed to generate HTML report"
    
    # Additional validation errors
    "ERR019" = "Mailbox contains items exceeding size limit"
    "ERR020" = "Unified Messaging is enabled and requires special handling"
    "ERR021" = "User belongs to nested groups that may require special handling"
    "ERR022" = "Mailbox has orphaned permissions"
    "ERR023" = "Mailbox has deeply nested folders that may cause issues"
    "ERR024" = "Mailbox has an excessive number of items"
    "ERR025" = "Shared calendar permissions require manual recreation"
    "ERR026" = "Arbitration or Audit mailbox requires special handling"
    
    # General errors
    "ERR999" = "Unspecified error during validation"
}

$script:MailboxResultProperties = @(
    # Basic mailbox information
    @{Name = "EmailAddress"; DefaultValue = $null},
    @{Name = "DisplayName"; DefaultValue = $null},
    @{Name = "UPN"; DefaultValue = $null},
    @{Name = "ExistingMailbox"; DefaultValue = $false},
    @{Name = "RecipientType"; DefaultValue = $null},
    @{Name = "RecipientTypeDetails"; DefaultValue = $null},
    @{Name = "MailboxEnabled"; DefaultValue = $false},
    
    # Domain and addressing properties
    @{Name = "UPNMatchesPrimarySMTP"; DefaultValue = $false},
    @{Name = "HasOnMicrosoftAddress"; DefaultValue = $false},
    @{Name = "HasRequiredOnMicrosoftAddress"; DefaultValue = $false},
    @{Name = "AllDomainsVerified"; DefaultValue = $false},
    @{Name = "HasUnverifiedDomains"; DefaultValue = $false},
    @{Name = "UnverifiedDomains"; DefaultValue = @()},
    @{Name = "ProxyAddressesFormatting"; DefaultValue = $true},
    
    # License information
    @{Name = "HasExchangeLicense"; DefaultValue = $false},
    @{Name = "HasE1OrE5License"; DefaultValue = $false},
    @{Name = "LicenseType"; DefaultValue = $null},
    @{Name = "LicenseProvisioningStatus"; DefaultValue = $null},
    @{Name = "LicenseDetails"; DefaultValue = $null},
    
    # Mailbox configuration
    @{Name = "HasSendAsPermissions"; DefaultValue = $false},
    @{Name = "SendAsPermissions"; DefaultValue = @()},
    @{Name = "HiddenFromAddressList"; DefaultValue = $false},
    @{Name = "SAMAccountNameValid"; DefaultValue = $false},
    @{Name = "SAMAccountName"; DefaultValue = $null},
    @{Name = "AliasValid"; DefaultValue = $false},
    @{Name = "Alias"; DefaultValue = $null},
    @{Name = "LitigationHoldEnabled"; DefaultValue = $false},
    @{Name = "RetentionHoldEnabled"; DefaultValue = $false},
    @{Name = "MailboxSizeGB"; DefaultValue = 0},
    @{Name = "MailboxSizeWarning"; DefaultValue = $false},
    @{Name = "ArchiveStatus"; DefaultValue = $null},
    @{Name = "ForwardingEnabled"; DefaultValue = $false},
    @{Name = "ForwardingAddress"; DefaultValue = $null},
    @{Name = "ExchangeGuid"; DefaultValue = $null},
    @{Name = "HasLegacyExchangeDN"; DefaultValue = $false},
    
    # Delegation and permissions
    @{Name = "FullAccessDelegates"; DefaultValue = @()},
    @{Name = "SendOnBehalfDelegates"; DefaultValue = @()},
    
    # Move request information
    @{Name = "PendingMoveRequest"; DefaultValue = $false},
    @{Name = "MoveRequestStatus"; DefaultValue = $null},
    
    # Extended validation properties
    @{Name = "UMEnabled"; DefaultValue = $false},
    @{Name = "UMDetails"; DefaultValue = $null},
    @{Name = "HasLargeItems"; DefaultValue = $false},
    @{Name = "LargeItemsDetails"; DefaultValue = @()},
    @{Name = "HasOrphanedPermissions"; DefaultValue = $false},
    @{Name = "OrphanedPermissions"; DefaultValue = @()},
    @{Name = "GroupMemberships"; DefaultValue = @()},
    @{Name = "DirectGroupCount"; DefaultValue = 0},
    @{Name = "NestedGroupCount"; DefaultValue = 0},
    @{Name = "FolderCount"; DefaultValue = 0},
    @{Name = "HasDeepFolderHierarchy"; DefaultValue = $false},
    @{Name = "DeepFolders"; DefaultValue = @()},
    @{Name = "HasLargeFolders"; DefaultValue = $false},
    @{Name = "LargeFolders"; DefaultValue = @()},
    @{Name = "TotalItemCount"; DefaultValue = 0},
    @{Name = "CalendarFolderCount"; DefaultValue = 0},
    @{Name = "CalendarItemCount"; DefaultValue = 0},
    @{Name = "HasSharedCalendars"; DefaultValue = $false},
    @{Name = "CalendarPermissions"; DefaultValue = @()},
    @{Name = "ContactFolderCount"; DefaultValue = 0},
    @{Name = "ContactItemCount"; DefaultValue = 0},
    @{Name = "IsArbitrationMailbox"; DefaultValue = $false},
    @{Name = "IsAuditLogMailbox"; DefaultValue = $false},
    
    # Result tracking
    @{Name = "Warnings"; DefaultValue = @()},
    @{Name = "Errors"; DefaultValue = @()},
    @{Name = "ErrorCodes"; DefaultValue = @()},
    @{Name = "OverallStatus"; DefaultValue = "Unknown"}
)

# Default configuration
$script:DefaultConfig = @{
    TargetDeliveryDomain = ""
    MigrationEndpointName = ""
    LogPath = ".\Logs\"
    ReportPath = ".\Reports\"
    NotificationEmails = @()
    CompleteAfterDays = 1
    MaxMailboxSizeGB = 50
    StartAfterMinutes = 15
    RequiredE1orE5License = $true
    CheckDomainVerification = $true
    BatchCreationTimeoutMinutes = 10
    HTMLTemplatePath = ".\Templates\ReportTemplate.html"
    RequireUPNMatchSMTP = $true
    TokenLifetimeMinutes = 50       # Default token lifetime in minutes
    TokenRefreshBuffer = 5          # Refresh tokens 5 minutes before expiration
    MaxItemSizeMB = 150             # Maximum size of items in MB
    ReportVersioning = $true        # Enable versioned report files
    DiskSpaceThresholdMB = 500      # Minimum disk space required in MB
    EnableVerboseLogging = $false   # Enable verbose logging
    LogCleanupDays = 30             # Number of days to keep logs
    ConnectionRetryCount = 3        # Number of retries for connection issues
    ConnectionRetryDelaySeconds = 5 # Delay between connection retries
    InactivityThresholdDays = 30    # Days of inactivity to consider a mailbox inactive
    MoveRequestDetailLevel = "Full" # Level of detail for move request reporting
    CheckIncompleteMailboxMoves = $true # Check for incomplete mailbox moves
    UseBadItemLimitRecommendations = $true # Use recommended BadItemLimit values
}

# Script-level variables
$script:Config = $null
$script:LogFile = $null
$script:BatchName = $null
$script:LogBuffer = @()
$script:StartTime = Get-Date
$script:ResourcesCreated = @()  # Track resources created for cleanup
#endregion Script Variables and Configuration

#region Functions - Initialization and Utilities

# Track authentication token lifetime
$global:LastConnectionTime = $null

function Write-Log {
    <#
    .SYNOPSIS
        Writes a log message to the console and log file.
    
    .DESCRIPTION
        Writes a formatted log message to the console and appends it to the log file.
        If the log file is not initialized yet, messages are buffered and written once the log file is available.
    
    .PARAMETER Message
        The message to log.
    
    .PARAMETER Level
        The severity level of the message (INFO, WARNING, ERROR, SUCCESS, CRITICAL, DEBUG).
    
    .PARAMETER ErrorCode
        An optional error code to include in the message.
    
    .PARAMETER Console
        If set to $true, the message is also written to the console. Default is $true.
    
    .EXAMPLE
        Write-Log -Message "Operation completed successfully" -Level "SUCCESS"
    
    .EXAMPLE
        Write-Log -Message "Failed to connect to server" -Level "ERROR" -ErrorCode "ERR004"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "CRITICAL", "DEBUG")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorCode,
        
        [Parameter(Mandatory = $false)]
        [switch]$Console = $true
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level]"
    
    # Add error code if specified
    if ($ErrorCode) {
        $logMessage += " [$ErrorCode]"
    }
    
    $logMessage += " $Message"
    
    # Set color based on level
    $color = switch ($Level) {
        "INFO" { "White" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        "SUCCESS" { "Green" }
        "CRITICAL" { "Magenta" }
        "DEBUG" { "Cyan" }
        default { "White" }
    }
    
    # Skip debug messages unless verbose logging is enabled
    if ($Level -eq "DEBUG" -and (-not $script:Config -or -not $script:Config.EnableVerboseLogging)) {
        return
    }
    
    # Write to console with color if requested
    if ($Console) {
        Write-Host $logMessage -ForegroundColor $color
    }
    
    # Append to log file if initialized
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logMessage
    }
    else {
        # If log file isn't initialized yet, buffer this message to write later
        $script:LogBuffer += $logMessage
    }
}

function Test-DiskSpace {
    <#
    .SYNOPSIS
        Checks if there is enough disk space available for the operation.
    
    .DESCRIPTION
        Verifies that the specified paths have enough disk space available
        before proceeding with operations that may require significant space.
    
    .PARAMETER Paths
        An array of paths to check for disk space.
    
    .PARAMETER ThresholdMB
        Minimum required disk space in MB.
    
    .EXAMPLE
        Test-DiskSpace -Paths @("C:\Logs", "C:\Reports") -ThresholdMB 500
    
    .OUTPUTS
        [bool] True if enough disk space is available, False otherwise.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Paths,
        
        [Parameter(Mandatory = $false)]
        [int]$ThresholdMB = 500
    )
    
    try {
        $allPathsHaveSpace = $true
        
        foreach ($path in $Paths) {
            # Get drive root for the path
            $drive = (Resolve-Path $path -ErrorAction SilentlyContinue | Split-Path -Qualifier) + "\"
            if (-not $drive) {
                # If path doesn't exist yet, get parent folder's drive
                $parentPath = Split-Path -Parent $path
                while ($parentPath -and -not (Test-Path $parentPath)) {
                    $parentPath = Split-Path -Parent $parentPath
                }
                
                if ($parentPath) {
                    $drive = (Resolve-Path $parentPath | Split-Path -Qualifier) + "\"
                }
                else {
                    # If still can't resolve, use current location
                    $drive = (Get-Location | Split-Path -Qualifier) + "\"
                }
            }
            
            # Get drive free space
            $driveInfo = Get-PSDrive -Name $drive[0] -PSProvider FileSystem
            $freeSpaceMB = [math]::Round($driveInfo.Free / 1MB, 2)
            
            Write-Log -Message "Free space on $drive is $freeSpaceMB MB (threshold: $ThresholdMB MB)" -Level "DEBUG"
            
            if ($freeSpaceMB -lt $ThresholdMB) {
                Write-Log -Message "Not enough disk space on $drive for $path. Available: $freeSpaceMB MB, Required: $ThresholdMB MB" -Level "ERROR"
                $allPathsHaveSpace = $false
            }
        }
        
        return $allPathsHaveSpace
    }
    catch {
        Write-Log -Message "Error checking disk space: $_" -Level "ERROR"
        return $false
    }
}

function Invoke-WithRetry {
    <#
    .SYNOPSIS
        Executes a script block with retry logic.
    
    .DESCRIPTION
        Attempts to execute a script block, retrying a specified number of times
        with a delay between attempts if the operation fails.
    
    .PARAMETER ScriptBlock
        The script block to execute.
    
    .PARAMETER MaxRetries
        The maximum number of retry attempts.
    
    .PARAMETER DelaySeconds
        The delay in seconds between retry attempts.
    
    .PARAMETER ExceptionMessage
        An optional message pattern to identify specific exceptions for retry.
        If specified, only errors matching this pattern will trigger a retry.
    
    .EXAMPLE
        Invoke-WithRetry -ScriptBlock { Get-Mailbox -Identity "user@domain.com" } -MaxRetries 3 -DelaySeconds 5
    
    .OUTPUTS
        Returns the output of the script block if successful.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 2,
        
        [Parameter(Mandatory = $false)]
        [int]$DelaySeconds = 3,
        
        [Parameter(Mandatory = $false)]
        [string]$ExceptionMessage = $null
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -le $MaxRetries) {
        try {
            if ($retryCount -gt 0) {
                Write-Log -Message "Retry $retryCount/$MaxRetries for operation..." -Level "DEBUG"
                Start-Sleep -Seconds $DelaySeconds
            }
            
            # Execute the script block
            $result = & $ScriptBlock
            $success = $true
        }
        catch {
            $errorMsg = $_.Exception.Message
            
            # Determine if we should retry based on exception message pattern
            $shouldRetry = $true
            if ($ExceptionMessage -and $errorMsg -notlike "*$ExceptionMessage*") {
                $shouldRetry = $false
            }
            
            $retryCount++
            
            if ($shouldRetry -and $retryCount -le $MaxRetries) {
                Write-Log -Message "Operation failed, will retry ($retryCount/$MaxRetries): $errorMsg" -Level "WARNING"
            }
            else {
                Write-Log -Message "Operation failed after $retryCount attempts: $errorMsg" -Level "ERROR"
                throw $_
            }
        }
    }
    
    return $result
}

function Test-TokenExpiration {
    <#
    .SYNOPSIS
        Checks if the authentication token is about to expire and refreshes if needed.
    
    .DESCRIPTION
        Monitors the age of authentication tokens and forces a refresh if they're
        approaching expiration to ensure continuous operation during long-running tasks.
    
    .PARAMETER ForceRefresh
        When specified, forces a token refresh regardless of the token age.
    
    .EXAMPLE
        Test-TokenExpiration
    
    .EXAMPLE
        Test-TokenExpiration -ForceRefresh
    
    .OUTPUTS
        [bool] True if token is valid or successfully refreshed, False otherwise.
    #>
    [CmdletBinding()]
    param(
        [switch]$ForceRefresh
    )
    
    try {
        # Check if token might be expiring soon
        if ($ForceRefresh -or 
            $global:LastConnectionTime -eq $null -or 
            ((Get-Date) - $global:LastConnectionTime).TotalMinutes -gt ($script:Config.TokenLifetimeMinutes - $script:Config.TokenRefreshBuffer)) {
            
            Write-Log -Message "Authentication token expiring soon or refresh requested - reconnecting services" -Level "INFO"
            return Connect-MigrationServices -ForceReconnect
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Error checking token expiration: $_" -Level "ERROR"
        return $false
    }
}

function Invoke-WithTokenRefresh {
    <#
    .SYNOPSIS
        Executes a script block with automatic token refresh on authentication failures.
    
    .DESCRIPTION
        Attempts to execute the script block and automatically refreshes the authentication
        token and retries if an authentication error occurs.
    
    .PARAMETER ScriptBlock
        The script block to execute.
    
    .PARAMETER MaxRetries
        Maximum number of retry attempts after token refresh.
    
    .EXAMPLE
        Invoke-WithTokenRefresh -ScriptBlock { Get-MigrationEndpoint -Identity "Endpoint1" }
    
    .OUTPUTS
        Returns the output of the script block if successful.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 1
    )
    
    $retryCount = 0
    
    while ($retryCount -le $MaxRetries) {
        try {
            return & $ScriptBlock
        }
        catch {
            $errorMessage = $_.Exception.Message
            
            # Check if this is an auth-related error
            if ($errorMessage -like "*token*" -or 
                $errorMessage -like "*unauthorized*" -or 
                $errorMessage -like "*authentication*" -or
                $errorMessage -like "*401*" -or 
                $errorMessage -like "*session*expired*") {
                
                $retryCount++
                
                if ($retryCount -le $MaxRetries) {
                    Write-Log -Message "Authentication error detected, refreshing token and retrying..." -Level "WARNING"
                    
                    if (Test-TokenExpiration -ForceRefresh) {
                        Write-Log -Message "Token refreshed successfully, retrying operation" -Level "INFO"
                        # Continue to next iteration to retry
                    }
                    else {
                        Write-Log -Message "Failed to refresh token, aborting operation" -Level "ERROR"
                        throw $_
                    }
                }
                else {
                    Write-Log -Message "Maximum retries reached after authentication errors" -Level "ERROR"
                    throw $_
                }
            }
            else {
                # Not an auth error, just rethrow
                throw $_
            }
        }
    }
}

function Test-Dependencies {
    <#
    .SYNOPSIS
        Checks if all required PowerShell modules are installed.
    
    .DESCRIPTION
        Verifies that all necessary PowerShell modules are installed with the required
        minimum versions and provides installation instructions if any are missing.
    
    .EXAMPLE
        Test-Dependencies
    
    .OUTPUTS
        [bool] True if all dependencies are met, False otherwise.
    #>
    [CmdletBinding()]
    param()
    
    Write-Log -Message "Checking for required PowerShell modules..." -Level "INFO"
    
    $requiredModules = @(
        @{
            Name = "ExchangeOnlineManagement"
            MinimumVersion = "3.0.0"
            Description = "Exchange Online PowerShell V3 module"
            InstallCommand = "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber"
        },
        @{
            Name = "Microsoft.Graph"
            MinimumVersion = "1.20.0" 
            Description = "Microsoft Graph PowerShell SDK"
            InstallCommand = "Install-Module -Name Microsoft.Graph -Force -AllowClobber"
        },
        @{
            Name = "Microsoft.Graph.Users"
            MinimumVersion = "1.20.0"
            Description = "Microsoft Graph Users module"
            InstallCommand = "Install-Module -Name Microsoft.Graph.Users -Force -AllowClobber"
        }
    )
    
    $missingModules = @()
    $outdatedModules = @()
    
    foreach ($module in $requiredModules) {
        $installedModule = Get-Module -Name $module.Name -ListAvailable
        
        if (-not $installedModule) {
            $missingModules += $module
            Write-Log -Message "Required module not found: $($module.Name) - $($module.Description)" -Level "ERROR" -ErrorCode "ERR001"
        }
        else {
            $latestVersion = ($installedModule | Sort-Object Version -Descending)[0].Version
            
            if ($latestVersion -lt [Version]$module.MinimumVersion) {
                $outdatedModules += @{
                    Module = $module
                    InstalledVersion = $latestVersion
                }
                Write-Log -Message "Module version too old: $($module.Name) - Current: $latestVersion, Required: $($module.MinimumVersion)" -Level "WARNING"
            }
            else {
                Write-Log -Message "Required module found: $($module.Name) (v$latestVersion)" -Level "DEBUG"
            }
        }
    }
    
    if ($missingModules.Count -gt 0 -or $outdatedModules.Count -gt 0) {
        Write-Log -Message "Dependency issues found. Please resolve before continuing." -Level "ERROR" -ErrorCode "ERR001"
        
        if ($missingModules.Count -gt 0) {
            Write-Log -Message "Missing modules:" -Level "ERROR"
            foreach ($module in $missingModules) {
                Write-Log -Message "- $($module.Name): $($module.Description)" -Level "ERROR"
                Write-Log -Message "  Install command: $($module.InstallCommand)" -Level "INFO"
            }
        }
        
        if ($outdatedModules.Count -gt 0) {
            Write-Log -Message "Outdated modules:" -Level "WARNING"
            foreach ($moduleInfo in $outdatedModules) {
                Write-Log -Message "- $($moduleInfo.Module.Name): Current v$($moduleInfo.InstalledVersion), Required v$($moduleInfo.Module.MinimumVersion)" -Level "WARNING"
                Write-Log -Message "  Update command: $($moduleInfo.Module.InstallCommand)" -Level "INFO"
            }
        }
        
        return $false
    }
    
    Write-Log -Message "All required dependencies are installed." -Level "SUCCESS"
    return $true
}

function Import-MigrationConfig {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        Write-Log -Message "Importing configuration from: $ConfigPath" -Level "INFO"
        
        # Check if config file exists
        if (-not (Test-Path -Path $ConfigPath)) {
            # If config file doesn't exist, create directory and default config
            $configDir = Split-Path -Path $ConfigPath -Parent
            
            if (-not (Test-Path -Path $configDir)) {
                New-Item -ItemType Directory -Path $configDir -Force | Out-Null
                Write-Log -Message "Created config directory: $configDir" -Level "INFO"
            }
            
            # Create default config file
            $script:DefaultConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $ConfigPath -Encoding utf8
            Write-Log -Message "Created default configuration file: $ConfigPath" -Level "INFO"
            
            # Return default config
            return $script:DefaultConfig
        }
        
        # Read and parse the JSON config file
        $configJson = Get-Content -Path $ConfigPath -Raw
        $config = $configJson | ConvertFrom-Json
        
        # Convert to hashtable
        $configHashtable = @{}
        
        # Get all default config properties and ensure they exist in loaded config
        foreach ($key in $script:DefaultConfig.Keys) {
            if (Get-Member -InputObject $config -Name $key -MemberType Properties) {
                $configHashtable[$key] = $config.$key
            }
            else {
                # Use default value if missing in config file
                $configHashtable[$key] = $script:DefaultConfig[$key]
                Write-Log -Message "Config file missing property '$key', using default value: $($script:DefaultConfig[$key])" -Level "WARNING"
            }
        }
        
        # Override with command-line parameters if specified
        if ($PSBoundParameters.ContainsKey('TargetDeliveryDomain') -and -not [string]::IsNullOrWhiteSpace($TargetDeliveryDomain)) {
            $configHashtable['TargetDeliveryDomain'] = $TargetDeliveryDomain
            Write-Log -Message "Using command-line override for TargetDeliveryDomain: $TargetDeliveryDomain" -Level "INFO"
        }
        
        if ($PSBoundParameters.ContainsKey('MigrationEndpointName') -and -not [string]::IsNullOrWhiteSpace($MigrationEndpointName)) {
            $configHashtable['MigrationEndpointName'] = $MigrationEndpointName
            Write-Log -Message "Using command-line override for MigrationEndpointName: $MigrationEndpointName" -Level "INFO"
        }
        
        if ($PSBoundParameters.ContainsKey('LogPath') -and -not [string]::IsNullOrWhiteSpace($LogPath)) {
            $configHashtable['LogPath'] = $LogPath
            Write-Log -Message "Using command-line override for LogPath: $LogPath" -Level "INFO"
        }
        
        if ($PSBoundParameters.ContainsKey('ReportPath') -and -not [string]::IsNullOrWhiteSpace($ReportPath)) {
            $configHashtable['ReportPath'] = $ReportPath
            Write-Log -Message "Using command-line override for ReportPath: $ReportPath" -Level "INFO"
        }
        
        if ($PSBoundParameters.ContainsKey('NotificationEmails') -and $NotificationEmails.Count -gt 0) {
            $configHashtable['NotificationEmails'] = $NotificationEmails
            Write-Log -Message "Using command-line override for NotificationEmails: $($NotificationEmails -join ', ')" -Level "INFO"
        }
        
        # Validate critical configuration items
        if ([string]::IsNullOrWhiteSpace($configHashtable['TargetDeliveryDomain'])) {
            Write-Log -Message "TargetDeliveryDomain is missing in config and not provided as parameter" -Level "ERROR" -ErrorCode "ERR002"
            return $null
        }
        
        if ([string]::IsNullOrWhiteSpace($configHashtable['MigrationEndpointName'])) {
            Write-Log -Message "MigrationEndpointName is missing in config and not provided as parameter" -Level "ERROR" -ErrorCode "ERR002"
            return $null
        }
        
        if ($configHashtable['NotificationEmails'].Count -eq 0) {
            Write-Log -Message "No NotificationEmails specified in config or as parameter" -Level "WARNING"
        }
        
        Write-Log -Message "Configuration loaded successfully." -Level "SUCCESS"
        return $configHashtable
    }
    catch {
        Write-Log -Message "Failed to load configuration: $_" -Level "ERROR" -ErrorCode "ERR002"
        return $null
    }
}

function Send-MigrationNotification {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        
        [Parameter(Mandatory = $true)]
        [string]$Body,
        
        [Parameter(Mandatory = $false)]
        [string[]]$To = $script:Config.NotificationEmails,
        
        [Parameter(Mandatory = $false)]
        [string]$From = "ExchangeMigration@$env:COMPUTERNAME",
        
        [Parameter(Mandatory = $false)]
        [string]$SmtpServer = "smtp.office365.com",
        
        [Parameter(Mandatory = $false)]
        [int]$Port = 587,
        
        [Parameter(Mandatory = $false)]
        [switch]$UseSsl = $true
    )
    
    if (-not $To -or $To.Count -eq 0) {
        Write-Log -Message "No notification email recipients specified" -Level "WARNING"
        return
    }
    
    try {
        Write-Log -Message "Sending email notification: $Subject" -Level "INFO"
        
        # Create a credential object if needed
        if (-not $script:EmailCredential) {
            $script:EmailCredential = Get-Credential -Message "Enter credentials for sending email notifications"
        }
        
        Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SmtpServer -Port $Port -UseSsl:$UseSsl -Credential $script:EmailCredential
        Write-Log -Message "Email notification sent successfully" -Level "SUCCESS"
    }
    catch {
        Write-Log -Message "Failed to send email notification: $_" -Level "ERROR"
    }
}

function Remove-OldLogs {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$LogCleanupDays = $script:Config.LogCleanupDays
    )
    
    try {
        $logPath = $script:Config.LogPath
        if (-not (Test-Path -Path $logPath)) {
            Write-Log -Message "Log directory does not exist: $logPath" -Level "WARNING"
            return
        }
        
        $cutoffDate = (Get-Date).AddDays(-$LogCleanupDays)
        $oldLogs = Get-ChildItem -Path $logPath -Filter "*.log" | Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        if ($oldLogs.Count -eq 0) {
            Write-Log -Message "No log files older than $LogCleanupDays days found" -Level "INFO"
            return
        }
        
        $removedCount = 0
        foreach ($log in $oldLogs) {
            try {
                # Check if file is in use before deleting
                $inUse = $false
                try {
                    $fileStream = [System.IO.File]::Open($log.FullName, 'Open', 'Read', 'None')
                    $fileStream.Close()
                    $fileStream.Dispose()
                }
                catch {
                    $inUse = $true
                }
                
                if (-not $inUse) {
                    Remove-Item -Path $log.FullName -Force
                    $removedCount++
                    Write-Log -Message "Removed old log file: $($log.Name)" -Level "DEBUG"
                }
                else {
                    Write-Log -Message "Skipping log file $($log.Name) as it is currently in use" -Level "DEBUG"
                }
            }
            catch {
                Write-Log -Message "Failed to remove log file $($log.Name): $_" -Level "WARNING"
            }
        }
        
        Write-Log -Message "Removed $removedCount of $($oldLogs.Count) log files older than $LogCleanupDays days" -Level "INFO"
    }
    catch {
        Write-Log -Message "Error during log cleanup: $_" -Level "ERROR"
    }
}

function Initialize-MigrationEnvironment {
    [CmdletBinding()]
    param()
    
    try {
        # Check disk space for all critical paths first
        $diskSpaceOK = Test-DiskSpace -Paths @(
            $script:Config.LogPath,
            $script:Config.ReportPath,
            (Split-Path -Path $BatchFilePath -Parent) # Include CSV folder
        ) -ThresholdMB $script:Config.DiskSpaceThresholdMB
        
        if (-not $diskSpaceOK) {
            Write-Log -Message "Insufficient disk space for operation. Migration aborted." -Level "ERROR"
            return $false
        }
        
        # Create log directory if it doesn't exist
        if (-not (Test-Path -Path $script:Config.LogPath)) {
            New-Item -ItemType Directory -Path $script:Config.LogPath -Force | Out-Null
            Write-Log -Message "Created log directory: $($script:Config.LogPath)" -Level "INFO"
        }
        
        # Create report directory if it doesn't exist
        if (-not (Test-Path -Path $script:Config.ReportPath)) {
            New-Item -ItemType Directory -Path $script:Config.ReportPath -Force | Out-Null
            Write-Log -Message "Created report directory: $($script:Config.ReportPath)" -Level "INFO"
        }
        
        # Create batch name from file name
        $script:BatchName = (Get-Item -Path $BatchFilePath).BaseName
        Write-Log -Message "Using batch name: $BatchName" -Level "INFO"
        
        # Initialize log file now that we have the batch name
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $script:LogFile = Join-Path -Path $script:Config.LogPath -ChildPath "$BatchName-$timestamp.log"
        
        # Write any buffered log messages to the log file
        if ($script:LogBuffer) {
            $script:LogBuffer | Out-File -FilePath $script:LogFile -Append
            $script:LogBuffer = $null
        }
        
        # Test CSV file
        if (-not (Test-Path -Path $BatchFilePath)) {
            Write-Log -Message "Batch file not found: $BatchFilePath" -Level "ERROR" -ErrorCode "ERR007"
            throw "Batch file not found: $BatchFilePath"
        }
        
        # Test CSV format
        $csvHeaders = (Get-Content -Path $BatchFilePath -TotalCount 1).Split(',')
        if ($csvHeaders -notcontains "EmailAddress") {
            Write-Log -Message "CSV file must contain 'EmailAddress' column" -Level "ERROR" -ErrorCode "ERR008"
            throw "Invalid CSV format: 'EmailAddress' column is required"
        }
        
        # Check if a migration batch with the same name already exists
        try {
            $existingBatch = Get-MigrationBatch -Identity $script:BatchName -ErrorAction SilentlyContinue
            if ($existingBatch) {
                Write-Log -Message "Migration batch with name '$script:BatchName' already exists with status: $($existingBatch.Status)" -Level "ERROR" -ErrorCode "ERR009"
                Write-Log -Message "Choose a different batch name by renaming your CSV file or remove the existing batch." -Level "ERROR"
                return $false
            }
        }
        catch {
            # If we get an error, it's likely because we're not connected yet, so we'll check this again later
            Write-Log -Message "Note: Will verify batch name availability after connecting to Exchange Online" -Level "INFO"
        }
        
        Write-Log -Message "Environment initialization completed successfully" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log -Message "Failed to initialize environment: $_" -Level "ERROR" -ErrorCode "ERR003"
        return $false
    }
}

function Connect-MigrationServices {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ForceReconnect,
        
        [Parameter(Mandatory = $false)]
        [int]$TokenLifetimeMinutes = 50,  # Token typically expires after 60 minutes
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 2
    )
    
    try {
        # Track if we need to reconnect based on token lifetime
        $global:LastConnectionTime = $global:LastConnectionTime ?? $null
        $tokenExpired = $false
        
        if ($global:LastConnectionTime -ne $null) {
            $timeSinceLastConnection = (Get-Date) - $global:LastConnectionTime
            $tokenExpired = $timeSinceLastConnection.TotalMinutes -gt $TokenLifetimeMinutes
            
            if ($tokenExpired) {
                Write-Log -Message "Authentication token may be expired (last connection: $global:LastConnectionTime)" -Level "WARNING"
            }
        }
        
        # Connect to Exchange Online if not already connected or if force reconnect is specified
        $exchangeConnected = $false
        $retryAttempt = 0
        
        while (-not $exchangeConnected -and $retryAttempt -lt $RetryCount) {
            try {
                if (-not $ForceReconnect -and -not $tokenExpired) {
                    # Check if already connected
                    $orgConfig = Get-OrganizationConfig -ErrorAction Stop
                    Write-Log -Message "Already connected to Exchange Online organization: $($orgConfig.DisplayName)" -Level "INFO"
                    $exchangeConnected = $true
                }
                else {
                    throw "Need to establish new connection"
                }
            }
            catch {
                $retryAttempt++
                Write-Log -Message "Connecting to Exchange Online (Attempt $retryAttempt of $RetryCount)..." -Level "INFO"
                
                try {
                    # Check for existing session and remove if necessary
                    $existingSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
                    if ($existingSession) {
                        Write-Log -Message "Removing existing Exchange Online session" -Level "INFO"
                        $existingSession | Remove-PSSession -ErrorAction SilentlyContinue
                    }
                    
                    # Use consistent modern authentication method regardless of PowerShell version
                    Connect-ExchangeOnline -ShowBanner:$false -UseDeviceAuthentication
                    
                    $orgConfig = Get-OrganizationConfig
                    $global:LastConnectionTime = Get-Date
                    $exchangeConnected = $true
                    
                    Write-Log -Message "Connected to Exchange Online organization: $($orgConfig.DisplayName)" -Level "SUCCESS"
                }
                catch {
                    if ($retryAttempt -ge $RetryCount) {
                        Write-Log -Message "Failed to connect to Exchange Online after $RetryCount attempts: $_" -Level "ERROR" -ErrorCode "ERR004"
                        Write-Log -Message "Check your internet connection and credentials. Try manually running Connect-ExchangeOnline in a separate window." -Level "ERROR"
                    }
                    else {
                        Write-Log -Message "Connection attempt failed, retrying in 5 seconds: $_" -Level "WARNING"
                        Start-Sleep -Seconds 5
                    }
                }
            }
        }
        
        # Connect to Microsoft Graph if not already connected
        $graphConnected = $false
        $retryAttempt = 0
        
        while (-not $graphConnected -and $retryAttempt -lt $RetryCount) {
            try {
                if (-not $ForceReconnect -and -not $tokenExpired) {
                    # Check if already connected
                    $mgContext = Get-MgContext -ErrorAction Stop
                    if ($mgContext) {
                        Write-Log -Message "Already connected to Microsoft Graph as: $($mgContext.Account)" -Level "INFO"
                        $graphConnected = $true
                    }
                    else {
                        throw "Not connected to Microsoft Graph"
                    }
                }
                else {
                    throw "Need to establish new connection"
                }
            }
            catch {
                $retryAttempt++
                Write-Log -Message "Connecting to Microsoft Graph (Attempt $retryAttempt of $RetryCount)..." -Level "INFO"
                
                try {
                    # Disconnect if there's an existing connection
                    if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
                        Disconnect-MgGraph -ErrorAction SilentlyContinue
                    }
    
                    # Use consistent modern authentication method regardless of PowerShell version
                    Connect-MgGraph -Scopes "Directory.Read.All", "User.Read.All" -UseDeviceAuthentication
    
                    $mgContext = Get-MgContext
                    $global:LastConnectionTime = Get-Date
                    $graphConnected = $true
    
                    Write-Log -Message "Connected to Microsoft Graph as: $($mgContext.Account)" -Level "SUCCESS"
                }
                catch [Microsoft.Graph.ServiceException] {
                    if ($retryAttempt -ge $RetryCount) {
                        Write-Log -Message "Microsoft Graph service error after $RetryCount attempts: $($_.Exception.Message)" -Level "ERROR" -ErrorCode "ERR005"
                        Write-Log -Message "Check your permissions and service health status" -Level "ERROR"
                    }
                    else {
                        Write-Log -Message "Microsoft Graph service error, retrying in 5 seconds: $($_.Exception.Message)" -Level "WARNING"
                        Start-Sleep -Seconds 5
                    }
                }
                catch [System.Net.WebException] {
                    if ($retryAttempt -ge $RetryCount) {
                        Write-Log -Message "Network error while connecting to Microsoft Graph after $RetryCount attempts: $($_.Exception.Message)" -Level "ERROR" -ErrorCode "ERR005"
                        Write-Log -Message "Check your network connectivity" -Level "ERROR"
                    }
                    else {
                        Write-Log -Message "Network error, retrying in 5 seconds: $($_.Exception.Message)" -Level "WARNING"
                        Start-Sleep -Seconds 5
                    }
                }
                catch {
                    if ($retryAttempt -ge $RetryCount) {
                        Write-Log -Message "Failed to connect to Microsoft Graph after $RetryCount attempts: $_" -Level "ERROR" -ErrorCode "ERR005"
                        Write-Log -Message "Check your internet connection and credentials. Try manually running Connect-MgGraph in a separate window." -Level "ERROR"
                    }
                    else {
                        Write-Log -Message "Connection attempt failed, retrying in 5 seconds: $_" -Level "WARNING"
                        Start-Sleep -Seconds 5
                    }
                }
            }
        }
        
        # Verify migration endpoint if both connections succeeded
        if ($exchangeConnected -and $graphConnected) {
            try {
                $migrationEndpoint = Get-MigrationEndpoint -Identity $script:Config.MigrationEndpointName -ErrorAction Stop
                Write-Log -Message "Migration endpoint validated: $($migrationEndpoint.Identity)" -Level "SUCCESS"
                
                # Check endpoint type to ensure it's appropriate for hybrid migration
                if ($migrationEndpoint.EndpointType -ne "ExchangeRemoteMove") {
                    Write-Log -Message "Warning: Migration endpoint type is $($migrationEndpoint.EndpointType), but ExchangeRemoteMove is recommended for hybrid migration" -Level "WARNING"
                }
                
                # Verify endpoint connectivity
                try {
                    $testResult = Test-MigrationServerAvailability -Endpoint $migrationEndpoint.Identity -ErrorAction Stop
                    Write-Log -Message "Migration endpoint connectivity test passed" -Level "SUCCESS"
                }
                catch {
                    Write-Log -Message "Migration endpoint connectivity test failed: $_" -Level "WARNING"
                    Write-Log -Message "The endpoint exists but may not be properly configured - check network connectivity" -Level "WARNING"
                }
            }
            catch {
                Write-Log -Message "Failed to validate migration endpoint: $_" -Level "ERROR" -ErrorCode "ERR006"
                Write-Log -Message "Available migration endpoints:" -Level "INFO"
                
                try {
                    $endpoints = Get-MigrationEndpoint
                    if ($endpoints.Count -eq 0) {
                        Write-Log -Message "  No migration endpoints found. You need to create one first." -Level "INFO"
                        Write-Log -Message "  Run 'New-MigrationEndpoint -ExchangeRemoteMove -Name \"Hybrid Migration Endpoint\" -RemoteServer [YourExchangeServerFQDN]'" -Level "INFO"
                    }
                    else {
                        foreach ($endpoint in $endpoints) {
                            Write-Log -Message "  - $($endpoint.Identity) ($($endpoint.EndpointType))" -Level "INFO"
                        }
                    }
                }
                catch {
                    Write-Log -Message "Could not retrieve available migration endpoints" -Level "WARNING"
                }
                
                return $false
            }
        }
        else {
            return $false
        }
        
        # Double-check if a migration batch with the same name already exists
        try {
            $existingBatch = Get-MigrationBatch -Identity $script:BatchName -ErrorAction SilentlyContinue
            if ($existingBatch) {
                Write-Log -Message "Migration batch with name '$script:BatchName' already exists with status: $($existingBatch.Status)" -Level "ERROR" -ErrorCode "ERR009"
                Write-Log -Message "Choose a different batch name by renaming your CSV file or remove the existing batch." -Level "ERROR"
                return $false
            }
        }
        catch {
            # If we get an error here, something is really wrong with the connection
            Write-Log -Message "Failed to check for existing migration batches: $_" -Level "ERROR"
            return $false
        }
        
        # Add a healthy connection check to ensure we have a good connection
        try {
            # Try a simple command to verify connection health
            $null = Get-AcceptedDomain -ErrorAction Stop
            $null = Get-MgUser -Top 1 -ErrorAction Stop
            
            Write-Log -Message "Connection health check passed for both Exchange Online and Microsoft Graph" -Level "SUCCESS"
            return $true
        }
        catch {
            Write-Log -Message "Connection health check failed: $_" -Level "ERROR"
            Write-Log -Message "Try reconnecting with -ForceReconnect parameter" -Level "ERROR"
            return $false
        }
    }
    catch {
        Write-Log -Message "Failed to connect to migration services: $_" -Level "ERROR"
        return $false
    }
}
#endregion Functions - Initialization and Utilities

#region Functions - Mailbox Validation

function Test-MailboxLicense {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        # Get user license info from Microsoft Graph
        $mgUser = Get-MgUser -UserId $EmailAddress -Property DisplayName, UserPrincipalName, UsageLocation, Id, ProxyAddresses -ErrorAction Stop
        $mgUserLicense = Get-MgUserLicenseDetail -UserId $EmailAddress -ErrorAction Stop
        
        # Modified license check: First check for E1 or E5 licenses
        $e1License = $mgUserLicense | Where-Object { 
            $_.SkuPartNumber -eq "STANDARDPACK" # E1
        }
        
        $e5License = $mgUserLicense | Where-Object { 
            $_.SkuPartNumber -eq "ENTERPRISEPACK" # E5
        }
        
        # Set license type based on what's found
        if ($e1License) {
            $Results.HasE1OrE5License = $true
            $Results.LicenseType = "E1"
            
            # Now check if Exchange is enabled within the E1 license
            $exchangeService = $e1License.ServicePlans | Where-Object {
                $_.ServicePlanName -eq "EXCHANGE_S_STANDARD"
            }
            
            if ($exchangeService) {
                $Results.HasExchangeLicense = $true
                $Results.LicenseDetails = "EXCHANGE_S_STANDARD"
                $Results.LicenseProvisioningStatus = $exchangeService.ProvisioningStatus
                
                if ($exchangeService.ProvisioningStatus -eq "Error") {
                    $Results.Errors += "E1 Exchange license provisioning error"
                    $Results.ErrorCodes += "ERR013"
                    Write-Log -Message "Error: E1 Exchange license provisioning error for $EmailAddress" -Level "ERROR" -ErrorCode "ERR013"
                    Write-Log -Message "Troubleshooting: Review the service health in Microsoft 365 admin center or try removing and reassigning the license." -Level "INFO"
                }
                elseif ($exchangeService.ProvisioningStatus -ne "Success") {
                    $Results.Warnings += "E1 Exchange license not fully provisioned: $($exchangeService.ProvisioningStatus)"
                    Write-Log -Message "Warning: E1 Exchange license not fully provisioned for $EmailAddress`: $($exchangeService.ProvisioningStatus)" -Level "WARNING"
                }
            }
            else {
                $Results.Errors += "E1 license found but Exchange Online service not enabled"
                $Results.ErrorCodes += "ERR012"
                Write-Log -Message "Error: E1 license found but Exchange Online service not enabled for $EmailAddress" -Level "ERROR" -ErrorCode "ERR012"
                Write-Log -Message "Troubleshooting: Verify the Exchange Online service is enabled in the license options" -Level "INFO"
            }
        }
        elseif ($e5License) {
            $Results.HasE1OrE5License = $true
            $Results.LicenseType = "E5"
            
            # Now check if Exchange is enabled within the E5 license
            $exchangeService = $e5License.ServicePlans | Where-Object {
                $_.ServicePlanName -eq "EXCHANGE_S_ENTERPRISE"
            }
            
            if ($exchangeService) {
                $Results.HasExchangeLicense = $true
                $Results.LicenseDetails = "EXCHANGE_S_ENTERPRISE"
                $Results.LicenseProvisioningStatus = $exchangeService.ProvisioningStatus
                
                if ($exchangeService.ProvisioningStatus -eq "Error") {
                    $Results.Errors += "E5 Exchange license provisioning error"
                    $Results.ErrorCodes += "ERR013"
                    Write-Log -Message "Error: E5 Exchange license provisioning error for $EmailAddress" -Level "ERROR" -ErrorCode "ERR013"
                    Write-Log -Message "Troubleshooting: Review the service health in Microsoft 365 admin center or try removing and reassigning the license." -Level "INFO"
                }
                elseif ($exchangeService.ProvisioningStatus -ne "Success") {
                    $Results.Warnings += "E5 Exchange license not fully provisioned: $($exchangeService.ProvisioningStatus)"
                    Write-Log -Message "Warning: E5 Exchange license not fully provisioned for $EmailAddress`: $($exchangeService.ProvisioningStatus)" -Level "WARNING"
                }
            }
            else {
                $Results.Errors += "E5 license found but Exchange Online service not enabled"
                $Results.ErrorCodes += "ERR012"
                Write-Log -Message "Error: E5 license found but Exchange Online service not enabled for $EmailAddress" -Level "ERROR" -ErrorCode "ERR012"
                Write-Log -Message "Troubleshooting: Verify the Exchange Online service is enabled in the license options" -Level "INFO"
            }
        }
        else {
            # Fallback to original check for any Exchange license if neither E1 nor E5 is found
            $mgUserExchangeLicense = $mgUserLicense.ServicePlans | Where-Object { 
                $_.ServicePlanName -like 'EXCHANGE_S_*' -and $_.AppliesTo -eq 'User'
            }
            
            if ($mgUserExchangeLicense) {
                $Results.HasExchangeLicense = $true
                $Results.LicenseDetails = $mgUserExchangeLicense.ServicePlanName -join ','
                $Results.LicenseProvisioningStatus = $mgUserExchangeLicense.ProvisioningStatus -join ','
                
                if ($script:Config.RequiredE1orE5License) {
                    $Results.Warnings += "Exchange license found but not E1 or E5: $($mgUserExchangeLicense.ServicePlanName -join ',')"
                    Write-Log -Message "Warning: Exchange license found but not E1 or E5 for $EmailAddress`: $($mgUserExchangeLicense.ServicePlanName -join ',')" -Level "WARNING"
                }
                
                if ($mgUserExchangeLicense.ProvisioningStatus -contains "Error") {
                    $Results.Errors += "License provisioning error"
                    $Results.ErrorCodes += "ERR013"
                    Write-Log -Message "Error: License provisioning error for $EmailAddress" -Level "ERROR" -ErrorCode "ERR013"
                    Write-Log -Message "Troubleshooting: Review the service health in Microsoft 365 admin center or try removing and reassigning the license." -Level "INFO"
                }
                elseif ($mgUserExchangeLicense.ProvisioningStatus -notcontains "Success") {
                    $Results.Warnings += "License not fully provisioned: $($mgUserExchangeLicense.ProvisioningStatus -join ',')"
                    Write-Log -Message "Warning: License not fully provisioned for $EmailAddress`: $($mgUserExchangeLicense.ProvisioningStatus -join ',')" -Level "WARNING"
                }
            }
            else {
                $Results.Errors += "No Exchange Online license assigned"
                $Results.ErrorCodes += "ERR012"
                Write-Log -Message "Error: No Exchange Online license assigned to $EmailAddress" -Level "ERROR" -ErrorCode "ERR012"
                Write-Log -Message "Troubleshooting: Assign an Exchange Online license (preferably E1 or E5) to the user" -Level "INFO"
            }
        }
        
        return $true
    }
    catch {
        $Results.Errors += "Failed to get user license info: $_"
        Write-Log -Message "Error: Failed to get user license info for $EmailAddress`: $_" -Level "ERROR"
        return $false
    }
}

function Test-MailboxConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        # Get Mailbox info from Exchange Online
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        $Results.ExistingMailbox = $true
        $Results.DisplayName = $mailbox.DisplayName
        $Results.RecipientType = $mailbox.RecipientType
        $Results.RecipientTypeDetails = $mailbox.RecipientTypeDetails
        $Results.MailboxEnabled = $true
        $Results.HiddenFromAddressList = $mailbox.HiddenFromAddressListsEnabled
        $Results.LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
        $Results.RetentionHoldEnabled = $mailbox.RetentionHoldEnabled
        $Results.ExchangeGuid = $mailbox.ExchangeGuid.ToString()
        $Results.HasLegacyExchangeDN = -not [string]::IsNullOrEmpty($mailbox.LegacyExchangeDN)
        $Results.ArchiveStatus = if ($mailbox.ArchiveStatus) { $mailbox.ArchiveStatus } else { "NotEnabled" }
        $Results.ForwardingEnabled = $mailbox.DeliverToMailboxAndForward
        $Results.ForwardingAddress = $mailbox.ForwardingAddress
        
        # Check for both litigation and retention hold enabled
        if ($mailbox.LitigationHoldEnabled -and $mailbox.RetentionHoldEnabled) {
            $Results.Warnings += "Both Litigation Hold and Retention Hold are enabled"
            Write-Log -Message "Warning: Both Litigation Hold and Retention Hold are enabled for $EmailAddress" -Level "WARNING"
        }
        
        # Check primary SMTP address
        $primarySMTP = ($mailbox.EmailAddresses | Where-Object { $_ -clike "SMTP:*" }).Substring(5)
        
        # Check UPN versus Primary SMTP
        if ($mailbox.UserPrincipalName -eq $primarySMTP) {
            $Results.UPNMatchesPrimarySMTP = $true
        }
        else {
            if ($script:Config.RequireUPNMatchSMTP) {
                $Results.Warnings += "UPN does not match primary SMTP address"
                Write-Log -Message "Warning: UPN ($($mailbox.UserPrincipalName)) does not match primary SMTP address ($primarySMTP) for $EmailAddress" -Level "WARNING"
                Write-Log -Message "Troubleshooting: Consider updating the UPN to match the primary SMTP address" -Level "INFO"
            }
        }
        
        $Results.UPN = $mailbox.UserPrincipalName
        
        # Check for onmicrosoft.com address
        if ($mailbox.EmailAddresses -like "*onmicrosoft.com*") {
            $Results.HasOnMicrosoftAddress = $true
        }
        else {
            $Results.Errors += "No onmicrosoft.com email address found"
            $Results.ErrorCodes += "ERR014"
            Write-Log -Message "Error: No onmicrosoft.com email address found for $EmailAddress" -Level "ERROR" -ErrorCode "ERR014"
            Write-Log -Message "Troubleshooting: Add an email alias with your tenant's onmicrosoft.com domain" -Level "INFO"
        }
        
        # Check proxy address formatting
        $primaryCount = ($mailbox.EmailAddresses | Where-Object { $_ -clike "SMTP:*" }).Count
        if ($primaryCount -ne 1) {
            $Results.ProxyAddressesFormatting = $false
            $Results.Errors += "Invalid number of primary SMTP addresses: $primaryCount"
            $Results.ErrorCodes += "ERR011"
            Write-Log -Message "Error: Invalid number of primary SMTP addresses ($primaryCount) for $EmailAddress" -Level "ERROR" -ErrorCode "ERR011"
            Write-Log -Message "Troubleshooting: Ensure exactly one primary SMTP address is set (capitalized SMTP:)" -Level "INFO"
        }
        
        # Check domain verification for all SMTP addresses if enabled in config
        if ($script:Config.CheckDomainVerification) {
            $allDomainsVerified = $true
            $acceptedDomains = Get-AcceptedDomain
            
            foreach ($address in $mailbox.EmailAddresses) {
                if ($address -like "smtp:*") {
                    $domain = ($address -split "@")[1]
                    if ($domain -notlike "*onmicrosoft.com") {
                        $domainVerified = $acceptedDomains | Where-Object { $_.DomainName -eq $domain -and $_.DomainType -eq "Authoritative" }
                        if (-not $domainVerified) {
                            $allDomainsVerified = $false
                            $Results.Errors += "Domain not verified: $domain"
                            $Results.ErrorCodes += "ERR015"
                            Write-Log -Message "Error: Domain not verified: $domain for $EmailAddress" -Level "ERROR" -ErrorCode "ERR015"
                            Write-Log -Message "Troubleshooting: Verify domain ownership in Microsoft 365 admin center" -Level "INFO"
                        }
                    }
                }
            }
            
            $Results.AllDomainsVerified = $allDomainsVerified
        }
        else {
            $Results.AllDomainsVerified = $true
        }
        
        return $true
    }
    catch {
        $Results.Errors += "Failed to get mailbox: $_"
        Write-Log -Message "Error: Failed to get mailbox for $EmailAddress`: $_" -Level "ERROR"
        return $false
    }
}

function Test-MailboxStatistics {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        # Get mailbox statistics
        $stats = Get-MailboxStatistics -Identity $EmailAddress
        $Results.MailboxSizeGB = [math]::Round($stats.TotalItemSize.Value.ToBytes() / 1GB, 2)
        
        if ($Results.MailboxSizeGB -gt $script:Config.MaxMailboxSizeGB) {
            $Results.MailboxSizeWarning = $true
            $Results.Warnings += "Mailbox size ($($Results.MailboxSizeGB) GB) exceeds warning threshold ($($script:Config.MaxMailboxSizeGB) GB)"
            Write-Log -Message "Warning: Mailbox size ($($Results.MailboxSizeGB) GB) exceeds warning threshold ($($script:Config.MaxMailboxSizeGB) GB) for $EmailAddress" -Level "WARNING"
            Write-Log -Message "Troubleshooting: Consider archiving older items or increasing the threshold in the configuration" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to get mailbox statistics: $_"
        Write-Log -Message "Warning: Failed to get mailbox statistics for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

function Get-MoveRequestDetails {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $false)]
        [string]$DetailLevel = $script:Config.MoveRequestDetailLevel
    )
    
    try {
        # Get move request with appropriate detail level
        switch ($DetailLevel) {
            "Full" {
                $moveRequest = Get-MoveRequestStatistics -Identity $EmailAddress -IncludeReport -DiagnosticInfo "Verbose" -ErrorAction SilentlyContinue
            }
            "Diagnostic" {
                $moveRequest = Get-MoveRequestStatistics -Identity $EmailAddress -IncludeReport -DiagnosticInfo "Verbose" -ErrorAction SilentlyContinue
            }
            "Basic" {
                $moveRequest = Get-MoveRequestStatistics -Identity $EmailAddress -ErrorAction SilentlyContinue
            }
            default {
                $moveRequest = Get-MoveRequestStatistics -Identity $EmailAddress -IncludeReport -ErrorAction SilentlyContinue
            }
        }
        
        return $moveRequest
    }
    catch {
        Write-Log -Message "Error getting move request details for $EmailAddress`: $_" -Level "WARNING"
        return $null
    }
}

function Test-MoveRequests {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results,
        
        [Parameter(Mandatory = $false)]
        [switch]$RemoveAndResubmitMoveRequests
    )
    
    try {
        # Check for pending move requests
        $moveRequest = Get-MoveRequest -Identity $EmailAddress -ErrorAction SilentlyContinue
        if ($moveRequest) {
            $Results.PendingMoveRequest = $true
            $Results.MoveRequestStatus = $moveRequest.Status
            
            # Get detailed move request information
            $moveRequestDetails = Get-MoveRequestDetails -EmailAddress $EmailAddress
            $Results | Add-Member -NotePropertyName "MoveRequestDetails" -NotePropertyValue $moveRequestDetails -Force
            
            $Results.Errors += "Pending move request found with status: $($moveRequest.Status)"
            $Results.ErrorCodes += "ERR016"
            Write-Log -Message "Error: Pending move request found with status: $($moveRequest.Status) for $EmailAddress" -Level "ERROR" -ErrorCode "ERR016"
            Write-Log -Message "Troubleshooting: Wait for the existing move request to complete or remove it with Remove-MoveRequest" -Level "INFO"
            
            # Handle remove and resubmit option
            if ($RemoveAndResubmitMoveRequests) {
                try {
                    Write-Log -Message "Removing existing move request for $EmailAddress" -Level "INFO"
                    Remove-MoveRequest -Identity $EmailAddress -Confirm:$false
                    Write-Log -Message "Existing move request removed successfully" -Level "SUCCESS"
                    
                    # Update results to reflect removal
                    $Results.PendingMoveRequest = $false
                    $Results.MoveRequestStatus = "Removed"
                    $Results.Errors = $Results.Errors | Where-Object { $_ -notlike "*Pending move request found*" }
                    $Results.ErrorCodes = $Results.ErrorCodes | Where-Object { $_ -ne "ERR016" }
                    
                    # Add a warning instead
                    $Results.Warnings += "Previous move request was removed and will be resubmitted"
                    Write-Log -Message "Mailbox will be included in new migration batch" -Level "INFO"
                }
                catch {
                    Write-Log -Message "Failed to remove move request for $EmailAddress`: $_" -Level "ERROR"
                    $Results.Warnings += "Failed to remove existing move request: $_"
                }
            }
        }
        
        # Check for orphaned move requests
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
        if ($mailbox) {
            $orphanedMoveRequests = Get-MoveRequest | Where-Object { 
                $_.DisplayName -eq $mailbox.DisplayName -and $_.Identity -ne $EmailAddress 
            }
            
            if ($orphanedMoveRequests) {
                foreach ($orphanedMove in $orphanedMoveRequests) {
                    $Results.Errors += "Orphaned move request found: $($orphanedMove.Identity) with status $($orphanedMove.Status)"
                    $Results.ErrorCodes += "ERR016"
                    Write-Log -Message "Error: Orphaned move request found: $($orphanedMove.Identity) with status $($orphanedMove.Status) for $EmailAddress" -Level "ERROR" -ErrorCode "ERR016"
                    Write-Log -Message "Troubleshooting: Remove orphaned move request with: Remove-MoveRequest -Identity '$($orphanedMove.Identity)'" -Level "INFO"
                    
                    # Handle remove and resubmit option for orphaned requests
                    if ($RemoveAndResubmitMoveRequests) {
                        try {
                            Write-Log -Message "Removing orphaned move request: $($orphanedMove.Identity)" -Level "INFO"
                            Remove-MoveRequest -Identity $orphanedMove.Identity -Confirm:$false
                            Write-Log -Message "Orphaned move request removed successfully" -Level "SUCCESS"
                            
                            # Add a warning
                            $Results.Warnings += "Orphaned move request was removed: $($orphanedMove.Identity)"
                        }
                        catch {
                            Write-Log -Message "Failed to remove orphaned move request $($orphanedMove.Identity): $_" -Level "ERROR"
                            $Results.Warnings += "Failed to remove orphaned move request: $_"
                        }
                    }
                }
            }
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check move requests: $_"
        Write-Log -Message "Warning: Failed to check move requests for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

function Test-MailboxPermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        # Check for delegates and permissions
        $fullAccessPermissions = Get-MailboxPermission -Identity $EmailAddress | Where-Object { 
            $_.AccessRights -contains "FullAccess" -and $_.User -notlike "NT AUTHORITY\*" -and $_.User -ne "Default"
        }
        
        if ($fullAccessPermissions) {
            $Results.FullAccessDelegates = $fullAccessPermissions | ForEach-Object { $_.User.ToString() }
            Write-Log -Message "Info: Full Access permissions found for $EmailAddress : $($Results.FullAccessDelegates -join ', ')" -Level "INFO"
        }
        
        # Check Send on Behalf permissions
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
        if ($mailbox -and $mailbox.GrantSendOnBehalfTo) {
            $Results.SendOnBehalfDelegates = $mailbox.GrantSendOnBehalfTo
            Write-Log -Message "Info: Send on Behalf permissions found for $EmailAddress : $($Results.SendOnBehalfDelegates -join ', ')" -Level "INFO"
        }
        
        # Check for Send As permissions
        $sendAsPermissions = Get-RecipientPermission -Identity $EmailAddress | Where-Object { 
            $_.AccessRights -contains "SendAs" -and $_.Trustee -notlike "NT AUTHORITY\*" -and $_.Trustee -ne "Default"
        }
        
        if ($sendAsPermissions) {
            $Results.HasSendAsPermissions = $true
            $Results.SendAsPermissions = $sendAsPermissions | ForEach-Object { $_.Trustee.ToString() }
            Write-Log -Message "Info: SendAs permissions found for $EmailAddress : $($sendAsPermissions.Trustee -join ', ')" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check mailbox permissions: $_"
        Write-Log -Message "Warning: Failed to check mailbox permissions for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
#endregion Functions - Mailbox Validation

#region Functions - Additional Validation

function Test-MailboxItemSizeLimits {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxItemSizeMB = 150
    )
    
    try {
        Write-Log -Message "Checking for large items in mailbox: $EmailAddress" -Level "INFO"
        
        # Get the mailbox
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Use Get-MailboxFolderStatistics to check for large items
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        $largeItems = @()
        
        foreach ($folder in $folderStats) {
            if ($folder.MaxItemSize -and [int64]::Parse(($folder.MaxItemSize -replace '[^\d]', '')) / 1MB -ge $MaxItemSizeMB) {
                $largeItems += [PSCustomObject]@{
                    FolderName = $folder.Name
                    FolderPath = $folder.FolderPath
                    MaxItemSize = $folder.MaxItemSize
                }
            }
        }
        
        if ($largeItems.Count -gt 0) {
            $Results.HasLargeItems = $true
            $Results.LargeItemsDetails = $largeItems
            $Results.Warnings += "Mailbox contains items larger than $MaxItemSizeMB MB, which may cause migration issues"
            
            Write-Log -Message "Warning: Found large items (>$MaxItemSizeMB MB) in mailbox $EmailAddress" -Level "WARNING"
            Write-Log -Message "Large items found in the following folders:" -Level "INFO"
            
            foreach ($item in $largeItems) {
                Write-Log -Message "  - Folder: $($item.FolderPath), Max item size: $($item.MaxItemSize)" -Level "INFO"
            }
            
            Write-Log -Message "Recommend: Review and remove/archive large items before migration" -Level "INFO"
        }
        else {
            $Results.HasLargeItems = $false
            Write-Log -Message "No items larger than $MaxItemSizeMB MB found in mailbox $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for large items: $_"
        Write-Log -Message "Warning: Failed to check for large items in mailbox $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

function Test-UnifiedMessagingConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking Unified Messaging configuration for: $EmailAddress" -Level "INFO"
        
        # Try to use Get-UMMailbox (if available in the environment)
        try {
            $umMailbox = Get-UMMailbox -Identity $EmailAddress -ErrorAction Stop
            if ($umMailbox) {
                $Results.UMEnabled = $true
                $Results.UMDetails = [PSCustomObject]@{
                    UMEnabled = $true
                    UMMailboxPolicy = $umMailbox.UMMailboxPolicy
                    SIPResourceIdentifier = $umMailbox.SIPResourceIdentifier
                    Extensions = $umMailbox.Extensions -join ", "
                }
                
                $Results.Warnings += "Unified Messaging is enabled and may require special handling during migration"
                Write-Log -Message "Warning: Unified Messaging is enabled for $EmailAddress" -Level "WARNING"
                Write-Log -Message "UM Policy: $($umMailbox.UMMailboxPolicy)" -Level "INFO"
                Write-Log -Message "Recommend: Disable UM before migration and plan to enable Cloud Voicemail post-migration" -Level "INFO"
            }
        }
        catch {
            # Explicitly set to null to prevent access to non-existent object
            $umMailbox = $null
            
            # Check recipient type as a fallback
            $recipient = Get-Recipient -Identity $EmailAddress -ErrorAction SilentlyContinue
            
            if ($recipient -and $recipient.RecipientTypeDetails -eq "UMEnabled") {
                $Results.UMEnabled = $true
                $Results.UMDetails = [PSCustomObject]@{
                    UMEnabled = $true
                    UMMailboxPolicy = "Unknown - use Exchange Admin Center to check"
                    Notes = "Detected via recipient type, use Exchange Admin Center for details"
                }
                
                $Results.Warnings += "Unified Messaging is enabled based on recipient type"
                Write-Log -Message "Warning: Unified Messaging appears to be enabled for $EmailAddress (detected via recipient type)" -Level "WARNING"
                Write-Log -Message "Recommend: Check Exchange Admin Center for UM details, disable UM before migration" -Level "INFO"
            }
            else {
                $Results.UMEnabled = $false
                Write-Log -Message "Unified Messaging is not enabled for $EmailAddress" -Level "INFO"
            }
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check Unified Messaging configuration: $_"
        Write-Log -Message "Warning: Failed to check Unified Messaging for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

function Test-OrphanedPermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking for orphaned permissions: $EmailAddress" -Level "INFO"
        
        # Get mailbox permissions
        $permissions = Get-MailboxPermission -Identity $EmailAddress | Where-Object { 
            $_.User -notlike "NT AUTHORITY\*" -and 
            $_.User -ne "Default" -and
            $_.IsInherited -eq $false
        }
        
        $orphanedPermissions = @()
        
        foreach ($permission in $permissions) {
            $user = $permission.User.ToString()
            
            # Skip if the user is a security principal (not a string representation)
            if ($permission.User -isnot [string]) {
                continue
            }
            
            # Check if the user exists
            try {
                $recipient = Get-Recipient -Identity $user -ErrorAction Stop
                # User exists, check if it's disabled
                if ($recipient.RecipientTypeDetails -eq "DisabledUser") {
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = $permission.AccessRights -join ", "
                        Issue = "User account is disabled"
                    }
                }
            }
            catch [Microsoft.Exchange.Management.Tasks.ManagementObjectNotFoundException] {
                # User not found in directory
                $orphanedPermissions += [PSCustomObject]@{
                    User = $user
                    AccessRights = $permission.AccessRights -join ", "
                    Issue = "User not found in directory"
                }
            }
            catch [System.Management.Automation.RemoteException] {
                # Handle remote exceptions (common with Exchange Online)
                if ($_.Exception.Message -like "*not found*") {
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = $permission.AccessRights -join ", "
                        Issue = "User not found in directory (remote error)"
                    }
                }
                else {
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = $permission.AccessRights -join ", "
                        Issue = "Error checking user: $($_.Exception.Message)"
                    }
                }
            }
            catch {
                # User might not exist anymore
                $orphanedPermissions += [PSCustomObject]@{
                    User = $user
                    AccessRights = $permission.AccessRights -join ", "
                    Issue = "User not found in directory"
                }
            }
        }
        
        # Check Send on Behalf permissions for orphaned users
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        if ($mailbox.GrantSendOnBehalfTo) {
            foreach ($user in $mailbox.GrantSendOnBehalfTo) {
                try {
                    $recipient = Get-Recipient -Identity $user -ErrorAction Stop
                    # User exists, check if it's disabled
                    if ($recipient.RecipientTypeDetails -eq "DisabledUser") {
                        $orphanedPermissions += [PSCustomObject]@{
                            User = $user
                            AccessRights = "SendOnBehalf"
                            Issue = "User account is disabled"
                        }
                    }
                }
                catch {
                    # User might not exist anymore
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = "SendOnBehalf"
                        Issue = "User not found in directory"
                    }
                }
            }
        }
        
        # Check Send As permissions for orphaned users
        $sendAsPermissions = Get-RecipientPermission -Identity $EmailAddress | Where-Object {
            $_.AccessRights -contains "SendAs" -and 
            $_.Trustee -notlike "NT AUTHORITY\*" -and 
            $_.Trustee -ne "Default"
        }
        
        foreach ($permission in $sendAsPermissions) {
            $user = $permission.Trustee.ToString()
            
            try {
                $recipient = Get-Recipient -Identity $user -ErrorAction Stop
                # User exists, check if it's disabled
                if ($recipient.RecipientTypeDetails -eq "DisabledUser") {
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = "SendAs"
                        Issue = "User account is disabled"
                    }
                }
            }
            catch {
                # User might not exist anymore
                $orphanedPermissions += [PSCustomObject]@{
                    User = $user
                    AccessRights = "SendAs"
                    Issue = "User not found in directory"
                }
            }
        }
        
        if ($orphanedPermissions.Count -gt 0) {
            $Results.HasOrphanedPermissions = $true
            $Results.OrphanedPermissions = $orphanedPermissions
            $Results.Warnings += "Mailbox has orphaned permissions that may not migrate correctly"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has orphaned permissions:" -Level "WARNING"
            foreach ($perm in $orphanedPermissions) {
                Write-Log -Message "  - User: $($perm.User), Rights: $($perm.AccessRights), Issue: $($perm.Issue)" -Level "WARNING"
            }
            Write-Log -Message "Recommend: Clean up orphaned permissions before migration" -Level "INFO"
        }
        else {
            $Results.HasOrphanedPermissions = $false
            Write-Log -Message "No orphaned permissions found for mailbox $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for orphaned permissions: $_"
        Write-Log -Message "Warning: Failed to check for orphaned permissions for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

function Get-NestedGroupMembership {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupIdentity,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxDepth = 5,
        
        [Parameter(Mandatory = $false)]
        [int]$CurrentDepth = 0,
        
        [Parameter(Mandatory = $false)]
        [System.Collections.ArrayList]$ProcessedGroups = $null
    )
    
    # Initialize the processed groups tracking array on first call
    if ($null -eq $ProcessedGroups) {
        $ProcessedGroups = New-Object System.Collections.ArrayList
    }
    
    # Prevent infinite recursion by stopping at max depth
    if ($CurrentDepth -ge $MaxDepth) {
        Write-Log -Message "Maximum recursion depth reached for group $GroupIdentity" -Level "WARNING"
        return @()
    }
    
    try {
        # Get the group
        $group = Get-Recipient -Identity $GroupIdentity -ErrorAction Stop
        
        # Add to processed groups to prevent circular references
        if ($ProcessedGroups -notcontains $group.DistinguishedName) {
            [void]$ProcessedGroups.Add($group.DistinguishedName)
        }
        else {
            # Already processed this group, skip to prevent circular reference
            return @()
        }
        
        # Get parent groups that contain this group
        $parentGroups = Get-Recipient -Filter "Members -eq '$($group.DistinguishedName)'" -RecipientTypeDetails 'GroupMailbox','MailUniversalDistributionGroup','MailUniversalSecurityGroup' -ErrorAction SilentlyContinue
        
        $results = @()
        foreach ($parent in $parentGroups) {
            $results += [PSCustomObject]@{
                GroupName = $parent.DisplayName
                GroupType = $parent.RecipientTypeDetails
                EmailAddress = $parent.PrimarySmtpAddress
                NestedLevel = $CurrentDepth + 1
                ParentOf = $group.DisplayName
            }
            
            # Recursively get each parent's parents
            $results += Get-NestedGroupMembership -GroupIdentity $parent.DistinguishedName -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1) -ProcessedGroups $ProcessedGroups
        }
        
        return $results
    }
    catch {
        Write-Log -Message "Error getting nested group membership for $GroupIdentity`: $_" -Level "WARNING"
        return @()
    }
}


function Test-RecursiveGroupMembership {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking recursive group membership: $EmailAddress" -Level "INFO"
        
        # Get mailbox
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Get direct group memberships
        $directGroups = Get-Recipient -Filter "Members -eq '$($mailbox.DistinguishedName)'" -RecipientTypeDetails 'GroupMailbox','MailUniversalDistributionGroup','MailUniversalSecurityGroup' -ErrorAction SilentlyContinue
        
        $groupInfo = @()
        $nestedGroups = @()
        
        if ($directGroups) {
            foreach ($group in $directGroups) {
                $groupInfo += [PSCustomObject]@{
                    GroupName = $group.DisplayName
                    GroupType = $group.RecipientTypeDetails
                    EmailAddress = $group.PrimarySmtpAddress
                    IsNested = $false
                    NestedLevel = 0
                }
                
                # Use the recursive function to get nested groups
                $nestedGroupMembers = Get-NestedGroupMembership -GroupIdentity $group.DistinguishedName
                
                if ($nestedGroupMembers -and $nestedGroupMembers.Count -gt 0) {
                    foreach ($nestedGroup in $nestedGroupMembers) {
                        $nestedGroups += [PSCustomObject]@{
                            GroupName = $nestedGroup.GroupName
                            GroupType = $nestedGroup.GroupType
                            EmailAddress = $nestedGroup.EmailAddress
                            NestedIn = $nestedGroup.ParentOf
                            IsNested = $true
                            NestedLevel = $nestedGroup.NestedLevel
                        }
                    }
                }
            }
        }
        
        # Combine direct and nested groups
        $allGroups = $groupInfo + $nestedGroups
        
        if ($allGroups.Count -gt 0) {
            $Results.GroupMemberships = $allGroups
            $Results.DirectGroupCount = ($allGroups | Where-Object { -not $_.IsNested }).Count
            $Results.NestedGroupCount = ($allGroups | Where-Object { $_.IsNested }).Count
            
            if ($nestedGroups.Count -gt 0) {
                $Results.Warnings += "User belongs to nested groups which may require special handling during migration"
                $maxNestedLevel = ($nestedGroups | Measure-Object -Property NestedLevel -Maximum).Maximum
                
                Write-Log -Message "Warning: User $EmailAddress belongs to nested groups (max depth: $maxNestedLevel):" -Level "WARNING"
                
                # Log only the first few nested groups to avoid excessive logging
                foreach ($group in ($nestedGroups | Select-Object -First 5)) {
                    Write-Log -Message "  - $($group.GroupName) ($($group.EmailAddress)) nested in $($group.NestedIn) (Level: $($group.NestedLevel))" -Level "WARNING"
                }
                
                if ($nestedGroups.Count -gt 5) {
                    Write-Log -Message "  - ... and $($nestedGroups.Count - 5) more nested groups" -Level "WARNING"
                }
                
                Write-Log -Message "Recommend: Document group memberships as they may need to be manually recreated" -Level "INFO"
            }
            else {
                Write-Log -Message "User $EmailAddress belongs to $($directGroups.Count) groups (no nested groups)" -Level "INFO"
            }
        }
        else {
            $Results.GroupMemberships = @()
            $Results.DirectGroupCount = 0
            $Results.NestedGroupCount = 0
            Write-Log -Message "No group memberships found for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check group memberships: $_"
        Write-Log -Message "Warning: Failed to check group memberships for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

function Test-MailboxFolderStructure {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking mailbox folder structure: $EmailAddress" -Level "INFO"
        
        # Get folder statistics
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        
        # Check folder count
        $folderCount = $folderStats.Count
        $Results.FolderCount = $folderCount
        
        # Check for deep folder hierarchy (more than 10 levels might cause issues)
        $deepFolders = $folderStats | Where-Object { ($_.FolderPath.Split('/').Count - 1) -gt 10 }
        
        if ($deepFolders) {
            $Results.HasDeepFolderHierarchy = $true
            $Results.DeepFolders = $deepFolders | Select-Object Name, FolderPath, ItemsInFolder
            $Results.Warnings += "Mailbox has deeply nested folders that might cause migration issues"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has deeply nested folders:" -Level "WARNING"
            foreach ($folder in $deepFolders | Select-Object -First 5) {
                $depth = $folder.FolderPath.Split('/').Count - 1
                Write-Log -Message "  - $($folder.FolderPath) (Depth: $depth, Items: $($folder.ItemsInFolder))" -Level "WARNING"
            }
            
            if ($deepFolders.Count -gt 5) {
                Write-Log -Message "  - ... and $($deepFolders.Count - 5) more deeply nested folders" -Level "WARNING"
            }
            
            Write-Log -Message "Recommend: Simplify folder structure before migration" -Level "INFO"
        }
        else {
            $Results.HasDeepFolderHierarchy = $false
            Write-Log -Message "No deeply nested folders found in mailbox $EmailAddress" -Level "INFO"
        }
        
        # Check for large folders (more than 5000 items might cause performance issues)
        $largeFolders = $folderStats | Where-Object { $_.ItemsInFolder -gt 5000 }
        
        if ($largeFolders) {
            $Results.HasLargeFolders = $true
            $Results.LargeFolders = $largeFolders | Select-Object Name, FolderPath, ItemsInFolder
            
            if ($largeFolders | Where-Object { $_.ItemsInFolder -gt 50000 }) {
                $Results.Warnings += "Mailbox has folders with more than 50,000 items, which might cause migration issues"
                Write-Log -Message "Warning: Mailbox $EmailAddress has folders with extremely high item counts:" -Level "WARNING"
            }
            else {
                Write-Log -Message "Note: Mailbox $EmailAddress has folders with high item counts:" -Level "INFO"
            }
            
            foreach ($folder in $largeFolders | Sort-Object ItemsInFolder -Descending | Select-Object -First 5) {
                Write-Log -Message "  - $($folder.FolderPath): $($folder.ItemsInFolder) items" -Level "INFO"
            }
            
            if ($largeFolders.Count -gt 5) {
                Write-Log -Message "  - ... and $($largeFolders.Count - 5) more large folders" -Level "INFO"
            }
        }
        else {
            $Results.HasLargeFolders = $false
            Write-Log -Message "No folders with high item counts found in mailbox $EmailAddress" -Level "INFO"
        }
        
        # Get total item count
        $totalItems = ($folderStats | Measure-Object -Property ItemsInFolder -Sum).Sum
        $Results.TotalItemCount = $totalItems
        
        if ($totalItems -gt 100000) {
            $Results.Warnings += "Mailbox has a very high total item count ($totalItems items), which might slow down migration"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a very high total item count: $totalItems items" -Level "WARNING"
            Write-Log -Message "Recommend: Consider archiving older items before migration to improve performance" -Level "INFO"
        }
        else {
            Write-Log -Message "Mailbox $EmailAddress has $totalItems items in total" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check mailbox folder structure: $_"
        Write-Log -Message "Warning: Failed to check mailbox folder structure for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

function Test-CalendarAndContactItems {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking calendar and contact items: $EmailAddress" -Level "INFO"
        
        # Get folder statistics
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        
        # Check calendar folders - using FolderType property which is locale-independent
        $calendarFolders = $folderStats | Where-Object { $_.FolderType -eq "Calendar" }
        $Results.CalendarFolderCount = $calendarFolders.Count
        $Results.CalendarItemCount = ($calendarFolders | Measure-Object -Property ItemsInFolder -Sum).Sum
        
        # Check for shared calendars
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        $calendarPermissions = @()
        
        # Get default calendar folder name (which may be localized)
        $defaultCalendarFolder = $calendarFolders | Where-Object { 
            # Check for default calendar paths in different languages
            $_.FolderPath -eq "/Calendar" -or 
            $_.FolderPath -eq "/Agenda" -or  # Dutch
            $_.IsDefaultFolder -eq $true
        }
        
        # Log detected default calendar folder for debugging
        if ($defaultCalendarFolder) {
            Write-Log -Message "Detected default calendar folder: $($defaultCalendarFolder.FolderPath)" -Level "DEBUG"
        }
        
        foreach ($calFolder in $calendarFolders) {
            # Use the path with escaped backslashes for PowerShell
            $folderPath = $calFolder.FolderPath.Replace('/', '\')
            
            # Handle different ways to construct folder identity
            try {
                # First try direct folder identity construction
                $folderIdentity = "$($mailbox.PrimarySmtpAddress):$folderPath"
                $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop
            }
            catch {
                try {
                    # If the first approach fails, try with the folder ID
                    $folderIdentity = "$($mailbox.PrimarySmtpAddress):\$($calFolder.FolderId)"
                    $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop
                }
                catch {
                    try {
                        # If that fails, try with the default format but removing the leading backslash
                        $folderIdentity = "$($mailbox.PrimarySmtpAddress):$($folderPath.TrimStart('\'))"
                        $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop
                    }
                    catch {
                        Write-Log -Message "Could not check permissions for calendar folder $($calFolder.FolderPath): $_" -Level "DEBUG"
                        continue
                    }
                }
            }
            
            # Process permissions if we got them
            $nonDefaultPermissions = $permissions | Where-Object { 
                $_.User.DisplayName -ne "Default" -and 
                $_.User.DisplayName -ne "Anonymous" -and
                $_.AccessRights -ne "None"
            }
            
            if ($nonDefaultPermissions) {
                foreach ($perm in $nonDefaultPermissions) {
                    $calendarPermissions += [PSCustomObject]@{
                        FolderPath = $calFolder.FolderPath
                        User = $perm.User.DisplayName
                        AccessRights = $perm.AccessRights -join ", "
                    }
                }
            }
        }
        
        if ($calendarPermissions.Count -gt 0) {
            $Results.HasSharedCalendars = $true
            $Results.CalendarPermissions = $calendarPermissions
            $Results.Warnings += "Mailbox has shared calendars with custom permissions that need to be recreated post-migration"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has shared calendars with custom permissions:" -Level "WARNING"
            foreach ($perm in $calendarPermissions | Select-Object -First 5) {
                Write-Log -Message "  - $($perm.FolderPath): Shared with $($perm.User) ($($perm.AccessRights))" -Level "WARNING"
            }
            
            if ($calendarPermissions.Count -gt 5) {
                Write-Log -Message "  - ... and $($calendarPermissions.Count - 5) more calendar sharing permissions" -Level "WARNING"
            }
            
            Write-Log -Message "Recommend: Document calendar sharing permissions as they may need to be manually recreated" -Level "INFO"
        }
        else {
            $Results.HasSharedCalendars = $false
            Write-Log -Message "No shared calendars found for mailbox $EmailAddress" -Level "INFO"
        }
        
        # Check for large number of recurring meetings
        if ($Results.CalendarItemCount -gt 1000) {
            $Results.Warnings += "Mailbox has a large number of calendar items ($($Results.CalendarItemCount)), which may include many recurring meetings"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a large number of calendar items: $($Results.CalendarItemCount)" -Level "WARNING"
            Write-Log -Message "Recommend: Consider cleaning up old calendar items before migration" -Level "INFO"
        }
        
        # Check contacts folders
        $contactFolders = $folderStats | Where-Object { $_.FolderType -eq "Contacts" }
        $Results.ContactFolderCount = $contactFolders.Count
        $Results.ContactItemCount = ($contactFolders | Measure-Object -Property ItemsInFolder -Sum).Sum
        
        if ($Results.ContactItemCount -gt 5000) {
            $Results.Warnings += "Mailbox has a large number of contacts ($($Results.ContactItemCount))"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a large number of contacts: $($Results.ContactItemCount)" -Level "WARNING"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check calendar and contact items: $_"
        Write-Log -Message "Warning: Failed to check calendar and contact items for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

function Test-ArbitrationMailboxes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        # Get mailbox
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Check if this is an arbitration mailbox
        if ($mailbox.RecipientTypeDetails -eq "ArbitrationMailbox") {
            $Results.IsArbitrationMailbox = $true
            $Results.Warnings += "This is an arbitration mailbox and requires special handling during migration"
            Write-Log -Message "Warning: $EmailAddress is an arbitration mailbox" -Level "WARNING"
            Write-Log -Message "Recommend: Follow special procedures for arbitration mailbox migration" -Level "INFO"
        }
        else {
            $Results.IsArbitrationMailbox = $false
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check if mailbox is an arbitration mailbox: $_"
        Write-Log -Message "Warning: Failed to check if $EmailAddress is an arbitration mailbox: $_" -Level "WARNING"
        return $false
    }
}

function Test-AuditLogMailboxes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        # Get mailbox
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Check if this is an audit log mailbox
        if ($mailbox.RecipientTypeDetails -eq "AuditLogMailbox") {
            $Results.IsAuditLogMailbox = $true
            $Results.Warnings += "This is an audit log mailbox and requires special handling during migration"
            Write-Log -Message "Warning: $EmailAddress is an audit log mailbox" -Level "WARNING"
            Write-Log -Message "Recommend: Follow special procedures for audit log mailbox migration" -Level "INFO"
        }
        else {
            $Results.IsAuditLogMailbox = $false
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check if mailbox is an audit log mailbox: $_"
        Write-Log -Message "Warning: Failed to check if $EmailAddress is an audit log mailbox: $_" -Level "WARNING"
        return $false
    }
}

# New function to create a mailbox result object with all properties
function New-MailboxTestResult {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )
    
    # Create a result object with all properties initialized
    $result = [PSCustomObject]@{}
    
    # Add each property with its default value
    foreach ($prop in $script:MailboxResultProperties) {
        # If it's the EmailAddress property, use the provided value
        if ($prop.Name -eq "EmailAddress") {
            $result | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $EmailAddress
        }
        else {
            $result | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.DefaultValue
        }
    }
    
    return $result
}

# Function to identify and check inactive mailboxes
function Test-InactiveMailboxes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results,
        
        [Parameter(Mandatory = $false)]
        [int]$InactivityThresholdDays = 30
    )
    
    try {
        Write-Log -Message "Checking for mailbox inactivity: $EmailAddress" -Level "INFO"
        
        # Get mailbox statistics to check last logon time
        $stats = Get-MailboxStatistics -Identity $EmailAddress -ErrorAction Stop
        
        # Add new properties to results for tracking inactive mailboxes
        $Results | Add-Member -NotePropertyName "LastLogonTime" -NotePropertyValue $stats.LastLogonTime -Force
        $Results | Add-Member -NotePropertyName "IsInactive" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "InactiveDays" -NotePropertyValue 0 -Force
        
        # Check if the mailbox has ever been logged into
        if ($null -eq $stats.LastLogonTime) {
            $Results.IsInactive = $true
            $Results.InactiveDays = 999 # Arbitrary high number to indicate "never used"
            $Results.Warnings += "Mailbox has never been logged into"
            Write-Log -Message "Warning: Mailbox $EmailAddress has never been logged into" -Level "WARNING"
        }
        else {
            # Calculate days since last logon
            $daysSinceLastLogon = (New-TimeSpan -Start $stats.LastLogonTime -End (Get-Date)).Days
            $Results.InactiveDays = $daysSinceLastLogon
            
            # Check if mailbox is considered inactive based on threshold
            if ($daysSinceLastLogon -gt $InactivityThresholdDays) {
                $Results.IsInactive = $true
                $Results.Warnings += "Mailbox inactive for $daysSinceLastLogon days (threshold: $InactivityThresholdDays days)"
                Write-Log -Message "Warning: Mailbox $EmailAddress inactive for $daysSinceLastLogon days" -Level "WARNING"
            }
            else {
                Write-Log -Message "Mailbox $EmailAddress is active (last logon: $($stats.LastLogonTime))" -Level "INFO"
            }
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check mailbox inactivity: $_"
        Write-Log -Message "Warning: Failed to check mailbox inactivity for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

# Function to check for potential corrupted items in a mailbox
function Test-MailboxCorruptItems {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking for potential corrupted items: $EmailAddress" -Level "INFO"
        
        # Add new properties to results
        $Results | Add-Member -NotePropertyName "RecommendedBadItemLimit" -NotePropertyValue 0 -Force
        $Results | Add-Member -NotePropertyName "PotentialCorruptionRisk" -NotePropertyValue "Low" -Force
        
        # Check mailbox item count - we'll use this to estimate potential for corruption
        # This is a heuristic approach since directly identifying corrupt items is challenging
        if ($Results.TotalItemCount -gt 100000) {
            $Results.PotentialCorruptionRisk = "High"
            # Recommend BadItemLimit as approximately 0.1% of total items, capped at 100
            $recommendedLimit = [Math]::Min(100, [Math]::Ceiling($Results.TotalItemCount * 0.001))
            $Results.RecommendedBadItemLimit = $recommendedLimit
            
            $Results.Warnings += "Large mailbox with high potential for corrupt items. Recommended BadItemLimit: $recommendedLimit"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a high risk of corrupt items due to size ($($Results.TotalItemCount) items)" -Level "WARNING"
            Write-Log -Message "Recommend BadItemLimit: $recommendedLimit for $EmailAddress" -Level "INFO"
        }
        elseif ($Results.TotalItemCount -gt 50000) {
            $Results.PotentialCorruptionRisk = "Medium"
            # Recommend BadItemLimit as approximately 0.05% of total items, capped at 50
            $recommendedLimit = [Math]::Min(50, [Math]::Ceiling($Results.TotalItemCount * 0.0005))
            $Results.RecommendedBadItemLimit = $recommendedLimit
            
            $Results.Warnings += "Medium-sized mailbox with potential for corrupt items. Recommended BadItemLimit: $recommendedLimit"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a medium risk of corrupt items ($($Results.TotalItemCount) items)" -Level "WARNING"
            Write-Log -Message "Recommend BadItemLimit: $recommendedLimit for $EmailAddress" -Level "INFO"
        }
        else {
            $Results.RecommendedBadItemLimit = 10 # Default conservative recommendation
            Write-Log -Message "Mailbox $EmailAddress has low risk of corrupt items" -Level "INFO"
        }
        
        # Check for deep folder paths which can sometimes indicate corruption
        if ($Results.HasDeepFolderHierarchy) {
            $Results.PotentialCorruptionRisk = "High"
            $Results.Warnings += "Deep folder hierarchy detected, which may indicate corrupted items"
            Write-Log -Message "Warning: Deep folder hierarchy in $EmailAddress may indicate corruption risk" -Level "WARNING"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for corrupted items: $_"
        Write-Log -Message "Warning: Failed to check for corrupted items in $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

# Function to check special mailbox types and provide specific guidance
function Test-SpecialMailboxTypes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking for special mailbox type: $EmailAddress" -Level "INFO"
        
        # Add new properties to results
        $Results | Add-Member -NotePropertyName "IsSpecialMailbox" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "SpecialMailboxType" -NotePropertyValue $null -Force
        $Results | Add-Member -NotePropertyName "SpecialMailboxGuidance" -NotePropertyValue $null -Force
        
        # Get mailbox to check type
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Define special mailbox types and their guidance
        $specialTypes = @{
            "SharedMailbox" = @{
                Guidance = "Ensure all users with access to this shared mailbox are migrated together. Permissions will need to be verified post-migration."
                Warning = "Shared mailbox detected - requires specific permission handling during migration"
            }
            "RoomMailbox" = @{
                Guidance = "Room mailboxes require calendar processing settings to be verified post-migration. Check booking policies and delegate access."
                Warning = "Room mailbox detected - requires calendar processing verification post-migration"
            }
            "EquipmentMailbox" = @{
                Guidance = "Equipment mailboxes require calendar processing settings to be verified post-migration. Check booking policies and delegate access."
                Warning = "Equipment mailbox detected - requires calendar processing verification post-migration"
            }
            "DiscoveryMailbox" = @{
                Guidance = "Discovery mailboxes should be handled separately from regular user migrations. Consider creating a new discovery mailbox in Exchange Online instead of migrating."
                Warning = "Discovery mailbox detected - not recommended for standard migration"
            }
            "MonitoringMailbox" = @{
                Guidance = "Monitoring mailboxes typically should not be migrated. Consider creating new monitoring configurations in Exchange Online."
                Warning = "Monitoring mailbox detected - not recommended for standard migration"
            }
            "LinkedMailbox" = @{
                Guidance = "Linked mailboxes must be converted to regular mailboxes before migration. This requires specific preparation steps."
                Warning = "Linked mailbox detected - requires conversion before migration"
            }
            "TeamMailbox" = @{
                Guidance = "Team mailboxes should be replaced with Microsoft 365 Groups instead of being directly migrated."
                Warning = "Team mailbox detected - consider using Microsoft 365 Groups instead"
            }
        }
        
        # Determine if this is a special mailbox type
        if ($specialTypes.ContainsKey($mailbox.RecipientTypeDetails)) {
            $Results.IsSpecialMailbox = $true
            $Results.SpecialMailboxType = $mailbox.RecipientTypeDetails
            $Results.SpecialMailboxGuidance = $specialTypes[$mailbox.RecipientTypeDetails].Guidance
            $Results.Warnings += $specialTypes[$mailbox.RecipientTypeDetails].Warning
            
            Write-Log -Message "Special mailbox type detected for $EmailAddress`: $($mailbox.RecipientTypeDetails)" -Level "WARNING"
            Write-Log -Message "Guidance: $($specialTypes[$mailbox.RecipientTypeDetails].Guidance)" -Level "INFO"
        }
        else {
            Write-Log -Message "Standard user mailbox detected for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for special mailbox type: $_"
        Write-Log -Message "Warning: Failed to check for special mailbox type for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

# Function to check and handle incomplete mailbox moves
function Test-IncompleteMailboxMoves {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking for incomplete mailbox moves: $EmailAddress" -Level "INFO"
        
        # Add new properties to results
        $Results | Add-Member -NotePropertyName "HasIncompleteMoves" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "IncompleteMovesDetails" -NotePropertyValue $null -Force
        
        # Check for completed but problematic move requests
        try {
            # Get completed move requests for this mailbox
            $completedMoves = Get-MoveRequestStatistics -Identity $EmailAddress -IncludeReport -ErrorAction SilentlyContinue | 
                Where-Object { $_.Status -eq "Completed" }
            
            if ($completedMoves) {
                # Check if there are any completed moves with failure reports
                $problemMoves = $completedMoves | Where-Object { 
                    $_.Report -like "*error*" -or 
                    $_.Report -like "*warning*" -or 
                    $_.Report -like "*failed*" -or
                    $_.BadItemsEncountered -gt 0 -or
                    $_.LargeItemsEncountered -gt 0
                }
                
                if ($problemMoves) {
                    $Results.HasIncompleteMoves = $true
                    $moveDetails = @()
                    
                    foreach ($move in $problemMoves) {
                        $moveDetails += [PSCustomObject]@{
                            CompletionTime = $move.CompletionTime
                            BadItems = $move.BadItemsEncountered
                            LargeItems = $move.LargeItemsEncountered
                            Errors = ($move.Report | Where-Object { $_ -like "*error*" }) -join "; "
                        }
                    }
                    
                    $Results.IncompleteMovesDetails = $moveDetails
                    $Results.Warnings += "Mailbox has previous migration attempts with issues"
                    
                    Write-Log -Message "Warning: Mailbox $EmailAddress has previous migration attempts with issues" -Level "WARNING"
                    Write-Log -Message "Found $($problemMoves.Count) move requests with potential problems" -Level "INFO"
                }
            }
        }
        catch {
            # It's ok if we can't get move statistics - might not have any
            Write-Log -Message "No existing move request statistics found for $EmailAddress" -Level "DEBUG"
        }
        
        # Check for mismatched Exchange GUIDs (less common issue)
        try {
            $adUser = Get-User -Identity $EmailAddress -ErrorAction SilentlyContinue
            
            if ($adUser -and $Results.ExchangeGuid -and $adUser.ExchangeGuid -ne $Results.ExchangeGuid) {
                $Results.HasIncompleteMoves = $true
                $Results.Warnings += "Mismatched Exchange GUIDs found - may indicate previous incomplete migration"
                
                Write-Log -Message "Warning: Mismatched Exchange GUIDs for $EmailAddress" -Level "WARNING"
                Write-Log -Message "AD Exchange GUID: $($adUser.ExchangeGuid), Mailbox Exchange GUID: $($Results.ExchangeGuid)" -Level "INFO"
            }
        }
        catch {
            Write-Log -Message "Unable to check for mismatched Exchange GUIDs: $_" -Level "DEBUG"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for incomplete mailbox moves: $_"
        Write-Log -Message "Warning: Failed to check for incomplete mailbox moves for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}


function Test-MailboxMigrationReadiness {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )
    
    # Use the new function to create a fully initialized result object
    $results = New-MailboxTestResult -EmailAddress $EmailAddress
    
    try {
        Write-Log -Message "Testing migration readiness for: $EmailAddress" -Level "INFO"
        
        # Call core test functions
        $mailboxConfigResult = Test-MailboxConfiguration -EmailAddress $EmailAddress -Results $results
        $licenseResult = Test-MailboxLicense -EmailAddress $EmailAddress -Results $results
        $statsResult = Test-MailboxStatistics -EmailAddress $EmailAddress -Results $results
        $moveRequestResult = Test-MoveRequests -EmailAddress $EmailAddress -Results $results
        $permissionsResult = Test-MailboxPermissions -EmailAddress $EmailAddress -Results $results
        
        # Call additional test functions
        $umResult = Test-UnifiedMessagingConfiguration -EmailAddress $EmailAddress -Results $results
        $itemSizeResult = Test-MailboxItemSizeLimits -EmailAddress $EmailAddress -Results $results
        $orphanedResult = Test-OrphanedPermissions -EmailAddress $EmailAddress -Results $results
        $groupResult = Test-RecursiveGroupMembership -EmailAddress $EmailAddress -Results $results
        $folderResult = Test-MailboxFolderStructure -EmailAddress $EmailAddress -Results $results
        $calendarResult = Test-CalendarAndContactItems -EmailAddress $EmailAddress -Results $results
        $arbitrationResult = Test-ArbitrationMailboxes -EmailAddress $EmailAddress -Results $results
        $auditResult = Test-AuditLogMailboxes -EmailAddress $EmailAddress -Results $results
        
        # Call new test functions for enhanced functionality
        $inactiveResult = Test-InactiveMailboxes -EmailAddress $EmailAddress -Results $results
        $corruptItemsResult = Test-MailboxCorruptItems -EmailAddress $EmailAddress -Results $results
        $specialTypeResult = Test-SpecialMailboxTypes -EmailAddress $EmailAddress -Results $results
        $incompleteMoveResult = Test-IncompleteMailboxMoves -EmailAddress $EmailAddress -Results $results
        
        # Determine overall status
        if ($results.Errors.Count -gt 0) {
            $results.OverallStatus = "Failed"
        }
        elseif ($results.Warnings.Count -gt 0) {
            $results.OverallStatus = "Warning"
        }
        else {
            $results.OverallStatus = "Ready"
        }
        
        return $results
    }
    catch {
        $results.Errors += "Test failed: $_"
        $results.OverallStatus = "Failed"
        Write-Log -Message "Test failed for $EmailAddress`: $_" -Level "ERROR"
        return $results
    }
}

function Export-HTMLReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$TestResults,
        
        [Parameter(Mandatory = $false)]
        [object]$MigrationBatch = $null
    )
    
    try {
        $reportPath = Join-Path -Path $script:Config.ReportPath -ChildPath "$script:BatchName-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
        
# Check for custom HTML template
$useCustomTemplate = $false
if (-not [string]::IsNullOrWhiteSpace($script:Config.HTMLTemplatePath) -and (Test-Path -Path $script:Config.HTMLTemplatePath)) {
    try {
        $htmlTemplate = Get-Content -Path $script:Config.HTMLTemplatePath -Raw
        $useCustomTemplate = $true
        Write-Log -Message "Using custom HTML template: $($script:Config.HTMLTemplatePath)" -Level "INFO"
    }
    catch {
        Write-Log -Message "Failed to load custom HTML template: $_" -Level "WARNING"
        $useCustomTemplate = $false
    }
}

if (-not $useCustomTemplate) {
    # Use default template
    $htmlTemplate = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Online Migration Report - $script:BatchName</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: #fff;
            padding: 20px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 5px;
        }
        h1, h2, h3 {
            color: #0078d4;
        }
        h1 {
            border-bottom: 2px solid #0078d4;
            padding-bottom: 10px;
            margin-top: 0;
        }
        .summary {
            background-color: #f0f8ff;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
            border-left: 5px solid #0078d4;
        }
        .status-ready {
            background-color: #dff0d8;
            color: #3c763d;
            padding: 5px 10px;
            border-radius: 3px;
            font-weight: bold;
        }
        .status-warning {
            background-color: #fcf8e3;
            color: #8a6d3b;
            padding: 5px 10px;
            border-radius: 3px;
            font-weight: bold;
        }
        .status-failed {
            background-color: #f2dede;
            color: #a94442;
            padding: 5px 10px;
            border-radius: 3px;
            font-weight: bold;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        th, td {
            text-align: left;
            padding: 12px 15px;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #0078d4;
            color: white;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .mailbox-details {
            margin-top: 20px;
            margin-bottom: 30px;
            border: 1px solid #ddd;
            border-radius: 5px;
            overflow: hidden;
        }
        .mailbox-details h3 {
            margin: 0;
            padding: 15px;
            background-color: #e9ecef;
            border-bottom: 1px solid #ddd;
        }
        .mailbox-details-content {
            padding: 15px;
        }
        .detail-row {
            display: flex;
            margin-bottom: 8px;
            border-bottom: 1px solid #eee;
            padding-bottom: 8px;
        }
        .detail-label {
            width: 30%;
            font-weight: bold;
        }
        .detail-value {
            width: 70%;
        }
        .error-list, .warning-list {
            margin-top: 10px;
            padding-left: 20px;
        }
        .error-list li {
            color: #a94442;
        }
        .warning-list li {
            color: #8a6d3b;
        }
        .badge {
            display: inline-block;
            min-width: 10px;
            padding: 3px 7px;
            font-size: 12px;
            font-weight: 700;
            line-height: 1;
            color: #fff;
            text-align: center;
            white-space: nowrap;
            vertical-align: middle;
            background-color: #777;
            border-radius: 10px;
        }
        .badge-success {
            background-color: #5cb85c;
        }
        .badge-warning {
            background-color: #f0ad4e;
        }
        .badge-danger {
            background-color: #d9534f;
        }
        .timestamp {
            color: #777;
            font-style: italic;
            margin-top: 30px;
            text-align: center;
        }
        .error-code {
            display: inline-block;
            background-color: #f8d7da;
            color: #721c24;
            padding: 0px 5px;
            border-radius: 3px;
            font-family: monospace;
            margin-right: 5px;
        }
        .action-required {
            background-color: #f8d7da;
            color: #721c24;
            padding: 10px;
            border-radius: 5px;
            margin-top: 20px;
            border-left: 5px solid #721c24;
        }
        .tabs {
            display: flex;
            margin-top: 20px;
            border-bottom: 1px solid #ddd;
        }
        .tab {
            padding: 10px 15px;
            cursor: pointer;
            background-color: #f1f1f1;
            border: 1px solid #ddd;
            border-bottom: none;
            margin-right: 5px;
            border-top-left-radius: 5px;
            border-top-right-radius: 5px;
            position: relative;
            top: 1px;
        }
        .tab.active {
            background-color: white;
            border-bottom: 1px solid white;
        }
        .tab-content {
            display: none;
        }
        .tab-content.active {
            display: block;
        }
        .category {
            margin-top: 15px;
            padding: 10px;
            background-color: #f8f9fa;
            border-radius: 3px;
            border-left: 3px solid #0078d4;
        }
        .category h4 {
            margin-top: 0;
            color: #0078d4;
        }
        .yes-value {
            color: #3c763d;
            font-weight: bold;
        }
        .no-value {
            color: #a94442;
            font-weight: bold;
        }
        .accordion {
            background-color: #f1f1f1;
            color: #444;
            cursor: pointer;
            padding: 10px;
            width: 100%;
            text-align: left;
            border: none;
            outline: none;
            transition: 0.4s;
            border-radius: 3px;
            margin-top: 5px;
        }
        .accordion:hover {
            background-color: #ddd;
        }
        .panel {
            padding: 0 18px;
            background-color: white;
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.2s ease-out;
        }
        .accordion:after {
            content: '\02795'; /* Unicode character for "plus" sign (+) */
            font-size: 10px;
            color: #777;
            float: right;
            margin-left: 5px;
        }
        .active:after {
            content: "\2796"; /* Unicode character for "minus" sign (-) */
        }
        .print-button {
            background-color: #0078d4;
            color: white;
            border: none;
            padding: 8px 15px;
            border-radius: 4px;
            cursor: pointer;
            margin-bottom: 15px;
        }
        .print-button:hover {
            background-color: #005a9e;
        }
        @media print {
            body {
                background-color: white;
                padding: 0;
                margin: 0;
            }
            .container {
                max-width: 100%;
                box-shadow: none;
                padding: 0;
            }
            .tabs, .tab, .print-button {
                display: none;
            }
            .tab-content {
                display: block;
            }
            .accordion, .panel {
                page-break-inside: avoid;
            }
            .accordion:after {
                display: none;
            }
            .panel {
                max-height: none !important;
                display: block !important;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <button class="print-button" onclick="window.print()">Print Report</button>
        <h1>Exchange Online Migration Report</h1>
        
        <div class="summary">
            <h2>Migration Batch: $script:BatchName</h2>
            <p><strong>Report Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p><strong>Total Mailboxes:</strong> $($TestResults.Count)</p>
            <p><strong>Ready for Migration:</strong> <span class="badge badge-success">$($TestResults | Where-Object { $_.OverallStatus -eq "Ready" } | Measure-Object).Count</span></p>
            <p><strong>Warnings:</strong> <span class="badge badge-warning">$($TestResults | Where-Object { $_.OverallStatus -eq "Warning" } | Measure-Object).Count</span></p>
            <p><strong>Failed:</strong> <span class="badge badge-danger">$($TestResults | Where-Object { $_.OverallStatus -eq "Failed" } | Measure-Object).Count</span></p>
            <p><strong>Inactive Mailboxes:</strong> <span class="badge badge-warning">$($TestResults | Where-Object { $_.IsInactive -eq $true } | Measure-Object).Count</span></p>
            <p><strong>Special Mailbox Types:</strong> <span class="badge badge-warning">$($TestResults | Where-Object { $_.IsSpecialMailbox -eq $true } | Measure-Object).Count</span></p>
            <p><strong>Script Version:</strong> $($script:ScriptVersion)</p>
"@

    # Add migration batch details if available
    if ($MigrationBatch) {
        $batchStatus = if ($MigrationBatch.IsDryRun) { "Dry Run - Not Created" } else { $MigrationBatch.Status }
        
        $htmlTemplate += @"
            <p><strong>Migration Batch Status:</strong> $batchStatus</p>
            <p><strong>Target Delivery Domain:</strong> $($script:Config.TargetDeliveryDomain)</p>
            <p><strong>Migration Endpoint:</strong> $($script:Config.MigrationEndpointName)</p>
"@
    }

    $htmlTemplate += @"
        </div>
        
        <div class="tabs">
            <div class="tab active" onclick="openTab(event, 'summary-tab')">Summary</div>
            <div class="tab" onclick="openTab(event, 'details-tab')">Detailed Results</div>
            <div class="tab" onclick="openTab(event, 'migration-tab')">Migration Guidance</div>
        </div>
        
        <div id="summary-tab" class="tab-content active">
            <h2>Mailbox Summary</h2>
            <table>
                <thead>
                    <tr>
                        <th>Email Address</th>
                        <th>Display Name</th>
                        <th>Mailbox Size (GB)</th>
                        <th>Items</th>
                        <th>Status</th>
                        <th>Issues</th>
                        <th>Special Type</th>
                        <th>Last Logon</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Add table rows for each mailbox
    foreach ($result in $TestResults) {
        $statusClass = switch ($result.OverallStatus) {
            "Ready" { "status-ready" }
            "Warning" { "status-warning" }
            "Failed" { "status-failed" }
            default { "" }
        }
        
        $issueCount = $result.Errors.Count + $result.Warnings.Count
        $specialType = if ($result.IsSpecialMailbox -eq $true) { $result.SpecialMailboxType } else { "Standard" }
        $lastLogon = if ($result.LastLogonTime) { (Get-Date $result.LastLogonTime -Format "yyyy-MM-dd") } else { "Never" }
        
        $htmlTemplate += @"
                <tr>
                    <td>$($result.EmailAddress)</td>
                    <td>$($result.DisplayName)</td>
                    <td>$($result.MailboxSizeGB)</td>
                    <td>$($result.TotalItemCount)</td>
                    <td><span class="$statusClass">$($result.OverallStatus)</span></td>
                    <td>$issueCount</td>
                    <td>$specialType</td>
                    <td>$lastLogon</td>
                </tr>
"@
    }

    $htmlTemplate += @"
                </tbody>
            </table>
        </div>
        
        <div id="details-tab" class="tab-content">
            <h2>Detailed Results</h2>
"@

    # Add detailed section for each mailbox
    foreach ($result in $TestResults) {
        $statusClass = switch ($result.OverallStatus) {
            "Ready" { "status-ready" }
            "Warning" { "status-warning" }
            "Failed" { "status-failed" }
            default { "" }
        }
        
        $htmlTemplate += @"
            <div class="mailbox-details">
                <h3>$($result.DisplayName) <small>($($result.EmailAddress))</small> - <span class="$statusClass">$($result.OverallStatus)</span></h3>
                <div class="mailbox-details-content">
                    <div class="detail-row">
                        <div class="detail-label">UPN</div>
                        <div class="detail-value">$($result.UPN)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Recipient Type</div>
                        <div class="detail-value">$($result.RecipientTypeDetails)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Mailbox Size</div>
                        <div class="detail-value">$($result.MailboxSizeGB) GB</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">UPN Matches Primary SMTP</div>
                        <div class="detail-value">$($result.UPNMatchesPrimarySMTP)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Has OnMicrosoft Address</div>
                        <div class="detail-value">$($result.HasOnMicrosoftAddress)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">All Domains Verified</div>
                        <div class="detail-value">$($result.AllDomainsVerified)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Has Exchange License</div>
                        <div class="detail-value">$($result.HasExchangeLicense)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">License Details</div>
                        <div class="detail-value">$($result.LicenseDetails)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">License Provisioning Status</div>
                        <div class="detail-value">$($result.LicenseProvisioningStatus)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Litigation Hold</div>
                        <div class="detail-value">$($result.LitigationHoldEnabled)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Retention Hold</div>
                        <div class="detail-value">$($result.RetentionHoldEnabled)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Archive Status</div>
                        <div class="detail-value">$($result.ArchiveStatus)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Forwarding Enabled</div>
                        <div class="detail-value">$($result.ForwardingEnabled)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Exchange GUID</div>
                        <div class="detail-value">$($result.ExchangeGuid)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Has Legacy Exchange DN</div>
                        <div class="detail-value">$($result.HasLegacyExchangeDN)</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-label">Pending Move Request</div>
                        <div class="detail-value">$($result.PendingMoveRequest)</div>
                    </div>
"@

        # Add move request status if applicable
        if ($result.PendingMoveRequest) {
            $htmlTemplate += @"
                    <div class="detail-row">
                        <div class="detail-label">Move Request Status</div>
                        <div class="detail-value">$($result.MoveRequestStatus)</div>
                    </div>
"@
        }

        # Add permissions if any
        if ($result.HasSendAsPermissions -or $result.FullAccessDelegates.Count -gt 0 -or $result.SendOnBehalfDelegates.Count -gt 0) {
            $htmlTemplate += @"
                    <div class="detail-row">
                        <div class="detail-label">Permissions</div>
                        <div class="detail-value">
"@
            
            if ($result.HasSendAsPermissions) {
                $htmlTemplate += @"
                            <strong>Send As:</strong> $($result.SendAsPermissions -join ", ")<br>
"@
            }
            
            if ($result.FullAccessDelegates.Count -gt 0) {
                $htmlTemplate += @"
                            <strong>Full Access:</strong> $($result.FullAccessDelegates -join ", ")<br>
"@
            }
            
            if ($result.SendOnBehalfDelegates.Count -gt 0) {
                $htmlTemplate += @"
                            <strong>Send On Behalf:</strong> $($result.SendOnBehalfDelegates -join ", ")
"@
            }
            
            $htmlTemplate += @"
                        </div>
                    </div>
"@
        }

        # Add special mailbox information
        $htmlTemplate += @"
                    <div class="category">
                        <h4>Special Mailbox Information</h4>
                        <div class="detail-row">
                            <div class="detail-label">Is Special Mailbox</div>
                            <div class="detail-value">$($result.IsSpecialMailbox)</div>
                        </div>
"@
        
        if ($result.IsSpecialMailbox) {
            $htmlTemplate += @"
                        <div class="detail-row">
                            <div class="detail-label">Special Mailbox Type</div>
                            <div class="detail-value">$($result.SpecialMailboxType)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Migration Guidance</div>
                            <div class="detail-value">$($result.SpecialMailboxGuidance)</div>
                        </div>
"@
        }
        
        $htmlTemplate += @"
                    </div>
                    
                    <div class="category">
                        <h4>Mailbox Activity Information</h4>
                        <div class="detail-row">
                            <div class="detail-label">Last Logon Time</div>
                            <div class="detail-value">$($result.LastLogonTime)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Is Inactive</div>
                            <div class="detail-value">$($result.IsInactive)</div>
                        </div>
"@
        
        if ($result.IsInactive) {
            $htmlTemplate += @"
                        <div class="detail-row">
                            <div class="detail-label">Inactive Days</div>
                            <div class="detail-value">$($result.InactiveDays)</div>
                        </div>
"@
        }
        
        $htmlTemplate += @"
                    </div>
                    
                    <div class="category">
                        <h4>Data Integrity Information</h4>
                        <div class="detail-row">
                            <div class="detail-label">Corruption Risk</div>
                            <div class="detail-value">$($result.PotentialCorruptionRisk)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Recommended BadItemLimit</div>
                            <div class="detail-value">$($result.RecommendedBadItemLimit)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Has Incomplete Moves</div>
                            <div class="detail-value">$($result.HasIncompleteMoves)</div>
                        </div>
"@

        if ($result.HasIncompleteMoves -and $result.IncompleteMovesDetails) {
            $htmlTemplate += @"
                        <div class="detail-row">
                            <div class="detail-label">Incomplete Moves Details</div>
                            <div class="detail-value">
                                <ul>
"@
            foreach ($moveDetail in $result.IncompleteMovesDetails) {
                $htmlTemplate += @"
                                    <li>Completion Time: $($moveDetail.CompletionTime), Bad Items: $($moveDetail.BadItems), Large Items: $($moveDetail.LargeItems)</li>
"@
            }
            $htmlTemplate += @"
                                </ul>
                            </div>
                        </div>
"@
        }

        $htmlTemplate += @"
                    </div>

                    <div class="category">
                        <h4>Extended Validation Results</h4>
                        <div class="detail-row">
                            <div class="detail-label">Total Items</div>
                            <div class="detail-value">$($result.TotalItemCount)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Folder Count</div>
                            <div class="detail-value">$($result.FolderCount)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Unified Messaging Enabled</div>
                            <div class="detail-value">$($result.UMEnabled)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Has Large Items (>150MB)</div>
                            <div class="detail-value">$($result.HasLargeItems)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Has Deeply Nested Folders</div>
                            <div class="detail-value">$($result.HasDeepFolderHierarchy)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Has Orphaned Permissions</div>
                            <div class="detail-value">$($result.HasOrphanedPermissions)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Security Group Memberships</div>
                            <div class="detail-value">Direct: $($result.DirectGroupCount), Nested: $($result.NestedGroupCount)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Calendar Items</div>
                            <div class="detail-value">$($result.CalendarItemCount) items in $($result.CalendarFolderCount) folders</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Has Shared Calendars</div>
                            <div class="detail-value">$($result.HasSharedCalendars)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Contact Items</div>
                            <div class="detail-value">$($result.ContactItemCount) items in $($result.ContactFolderCount) folders</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Is Arbitration Mailbox</div>
                            <div class="detail-value">$($result.IsArbitrationMailbox)</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-label">Is Audit Log Mailbox</div>
                            <div class="detail-value">$($result.IsAuditLogMailbox)</div>
                        </div>
                    </div>
"@

        # Show large items if any exist
        if ($result.HasLargeItems -and $result.LargeItemsDetails.Count -gt 0) {
            $htmlTemplate += @"
                    <div class="category">
                        <h4>Large Items Details</h4>
                        <table>
                            <thead>
                                <tr>
                                    <th>Folder Path</th>
                                    <th>Max Item Size</th>
                                </tr>
                            </thead>
                            <tbody>
"@
            foreach ($item in $result.LargeItemsDetails) {
                $htmlTemplate += @"
                                <tr>
                                    <td>$($item.FolderPath)</td>
                                    <td>$($item.MaxItemSize)</td>
                                </tr>
"@
            }
            $htmlTemplate += @"
                            </tbody>
                        </table>
                    </div>
"@
        }

        # Add errors and warnings
        if ($result.Errors.Count -gt 0 -or $result.Warnings.Count -gt 0) {
            $htmlTemplate += @"
                    <div class="detail-row">
                        <div class="detail-label">Issues</div>
                        <div class="detail-value">
"@
            
            if ($result.Errors.Count -gt 0) {
                $htmlTemplate += @"
                            <strong>Errors:</strong>
                            <ul class="error-list">
"@
                
                for ($i = 0; $i -lt $result.Errors.Count; $i++) {
                    $errorCode = if ($i -lt $result.ErrorCodes.Count) { $result.ErrorCodes[$i] } else { "" }
                    $errorCodeHtml = if ($errorCode) { "<span class='error-code'>$errorCode</span>" } else { "" }
                    
                    $htmlTemplate += @"
                                <li>$errorCodeHtml$($result.Errors[$i])</li>
"@
                }
                
                $htmlTemplate += @"
                            </ul>
"@
            }
            
            if ($result.Warnings.Count -gt 0) {
                $htmlTemplate += @"
                            <strong>Warnings:</strong>
                            <ul class="warning-list">
"@
                
                foreach ($warning in $result.Warnings) {
                    $htmlTemplate += @"
                                <li>$warning</li>
"@
                }
                
                $htmlTemplate += @"
                            </ul>
"@
            }
            
            $htmlTemplate += @"
                        </div>
                    </div>
"@
        }

        $htmlTemplate += @"
                </div>
            </div>
"@
    }

    # Add migration guidance tab
    $htmlTemplate += @"
        </div>
        
        <div id="migration-tab" class="tab-content">
            <h2>Migration Guidance</h2>
            
            <div class="category">
                <h4>Critical Issues</h4>
                <p>The following issues must be resolved before migration:</p>
                <ul id="critical-issues">
"@

    # Add critical issues
    $criticalIssues = @()
    foreach ($result in $TestResults) {
        if ($result.Errors.Count -gt 0) {
            $criticalIssues += "<li><strong>$($result.DisplayName) ($($result.EmailAddress))</strong>: $($result.Errors[0])</li>"
        }
    }

    if ($criticalIssues.Count -gt 0) {
        $htmlTemplate += $criticalIssues -join "`n"
    }
    else {
        $htmlTemplate += "<li>No critical issues found. All mailboxes can be migrated.</li>"
    }

    $htmlTemplate += @"
                </ul>
            </div>
            
            <div class="category">
                <h4>Performance Considerations</h4>
                <p>The following issues may impact migration performance:</p>
                <ul id="performance-issues">
"@

    # Add performance considerations
    $performanceIssues = @()
    foreach ($result in $TestResults) {
        if ($result.MailboxSizeGB -gt $script:Config.MaxMailboxSizeGB) {
            $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: Large mailbox size ($($result.MailboxSizeGB) GB)</li>"
        }
        if ($result.HasLargeFolders) {
            $folderCount = ($result.LargeFolders | Where-Object { $_.ItemsInFolder -gt 50000 }).Count
if ($folderCount -gt 0) {
                $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: Has $folderCount folders with over 50,000 items</li>"
            }
        }
        if ($result.TotalItemCount -gt 100000) {
            $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: High item count ($($result.TotalItemCount) items)</li>"
        }
        if ($result.HasLargeItems) {
            $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: Contains items larger than 150 MB</li>"
        }
        if ($result.PotentialCorruptionRisk -eq "High") {
            $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: High risk of corrupted items (Recommended BadItemLimit: $($result.RecommendedBadItemLimit))</li>"
        }
    }

    if ($performanceIssues.Count -gt 0) {
        $htmlTemplate += $performanceIssues -join "`n"
    }
    else {
        $htmlTemplate += "<li>No performance issues found.</li>"
    }

    $htmlTemplate += @"
                </ul>
            </div>

            <div class="category">
                <h4>Special Mailbox Types</h4>
                <p>The following special mailbox types require specific handling:</p>
                <ul id="special-mailboxes">
"@

    # Add special mailboxes
    $specialMailboxes = @()
    foreach ($result in $TestResults) {
        if ($result.IsSpecialMailbox) {
            $specialMailboxes += "<li><strong>$($result.DisplayName) ($($result.EmailAddress))</strong>: $($result.SpecialMailboxType) - $($result.SpecialMailboxGuidance)</li>"
        }
    }

    if ($specialMailboxes.Count -gt 0) {
        $htmlTemplate += $specialMailboxes -join "`n"
    }
    else {
        $htmlTemplate += "<li>No special mailbox types found.</li>"
    }

    $htmlTemplate += @"
                </ul>
            </div>
            
            <div class="category">
                <h4>Post-Migration Tasks</h4>
                <p>The following tasks should be performed after migration:</p>
                <ul id="post-migration-tasks">
"@

    # Add post-migration tasks
    $postMigrationTasks = @()
    foreach ($result in $TestResults) {
        if ($result.UMEnabled) {
            $postMigrationTasks += "<li><strong>$($result.DisplayName)</strong>: Configure Cloud Voicemail to replace Unified Messaging</li>"
        }
        if ($result.HasSharedCalendars) {
            $postMigrationTasks += "<li><strong>$($result.DisplayName)</strong>: Recreate calendar sharing permissions</li>"
        }
        if ($result.NestedGroupCount -gt 0) {
            $postMigrationTasks += "<li><strong>$($result.DisplayName)</strong>: Verify nested group memberships</li>"
        }
        if ($result.HasIncompleteMoves) {
            $postMigrationTasks += "<li><strong>$($result.DisplayName)</strong>: Verify mailbox content completeness due to previous incomplete migration</li>"
        }
    }

    if ($postMigrationTasks.Count -gt 0) {
        $htmlTemplate += $postMigrationTasks -join "`n"
    }
    else {
        $htmlTemplate += "<li>No specific post-migration tasks identified.</li>"
    }

    $htmlTemplate += @"
                </ul>
            </div>
            
            <div class="category">
                <h4>Special Mailbox Migration Guidance</h4>
                <button class="accordion">Handling Special Mailbox Types</button>
                <div class="panel">
                    <ul>
                        <li><strong>Shared Mailboxes</strong>: Ensure all users with access are migrated together. Verify permissions post-migration.</li>
                        <li><strong>Resource Mailboxes</strong>: Verify calendar processing settings after migration.</li>
                        <li><strong>Discovery/Monitoring Mailboxes</strong>: Consider creating new ones in Exchange Online rather than migrating.</li>
                        <li><strong>Team Mailboxes</strong>: Consider replacing with Microsoft 365 Groups instead of direct migration.</li>
                    </ul>
                </div>
                
                <button class="accordion">Dealing with Corrupted Items</button>
                <div class="panel">
                    <ul>
                        <li>Use BadItemLimit parameter when necessary for large mailboxes with potential corruption.</li>
                        <li>Example: <code>New-MoveRequest -Identity 'User' -Remote -RemoteHostName 'webmail.contoso.com' -RemoteCredential $onpremCred -TargetDeliveryDomain 'contoso.mail.onmicrosoft.com' -BadItemLimit 40</code></li>
                        <li>Be aware that using BadItemLimit may result in data loss of corrupted items.</li>
                        <li>Documentation: <a href="#" onclick="return false;">Troubleshoot migration issues in Exchange hybrid</a></li>
                    </ul>
                </div>
                
                <button class="accordion">Handling Incomplete Migrations</button>
                <div class="panel">
                    <ul>
                        <li>Mailboxes may show as completed in Exchange Online but still appear as on-premises "User Mailbox" instead of "Remote Mailbox."</li>
                        <li>Common causes: internet connectivity breaks, user password changes during migration, or MRS proxy timeouts.</li>
                        <li>Check if mailboxes have conflicting Exchange GUIDs between on-premises and cloud.</li>
                        <li>Remove any orphaned move requests before attempting a new migration.</li>
                    </ul>
                </div>
            </div>
            
            <div class="category">
                <h4>Migration Recommendations</h4>
                <button class="accordion">General Migration Best Practices</button>
                <div class="panel">
                    <ul>
                        <li>Test migrations with pilot users before full deployment</li>
                        <li>Schedule migrations during off-hours to minimize disruption</li>
                        <li>Communicate with users before, during, and after migration</li>
                        <li>Have a rollback plan in case of issues</li>
                        <li>Monitor migration batches closely for failures</li>
                    </ul>
                </div>
                
                <button class="accordion">Handling Large Mailboxes</button>
                <div class="panel">
                    <ul>
                        <li>Consider using multiple smaller batches for large mailboxes</li>
                        <li>Encourage users to clean up mailboxes before migration</li>
                        <li>Archive older content before migration</li>
                        <li>Monitor network bandwidth during migration</li>
                    </ul>
                </div>
                
                <button class="accordion">User Experience Considerations</button>
                <div class="panel">
                    <ul>
                        <li>Provide clear instructions for accessing new mailboxes</li>
                        <li>Reset Outlook profiles after migration if needed</li>
                        <li>Test mobile device access</li>
                        <li>Provide training on any new features in Exchange Online</li>
                    </ul>
                </div>
            </div>
        </div>
"@

    # Add action required section
    $failedMailboxes = $TestResults | Where-Object { $_.OverallStatus -eq "Failed" }
    if ($failedMailboxes.Count -gt 0) {
        $htmlTemplate += @"
            <div class="action-required">
                <h3>Action Required</h3>
                <p>The following mailboxes have issues that need to be resolved before migration:</p>
                <ul>
"@
        
        foreach ($mailbox in $failedMailboxes) {
            $htmlTemplate += @"
                    <li><strong>$($mailbox.DisplayName) ($($mailbox.EmailAddress))</strong> - Issues: $($mailbox.Errors.Count)</li>
"@
        }
        
        $htmlTemplate += @"
                </ul>
            </div>
"@
    }

    # Close HTML
    $htmlTemplate += @"
        
        <p class="timestamp">Generated by Exchange Online Migration Script v$($script:ScriptVersion)</p>
    </div>
    
    <script>
    function openTab(evt, tabName) {
        var i, tabContent, tabLinks;
        
        tabContent = document.getElementsByClassName("tab-content");
        for (i = 0; i < tabContent.length; i++) {
            tabContent[i].className = tabContent[i].className.replace(" active", "");
        }
        
        tabLinks = document.getElementsByClassName("tab");
        for (i = 0; i < tabLinks.length; i++) {
            tabLinks[i].className = tabLinks[i].className.replace(" active", "");
        }
        
        document.getElementById(tabName).className += " active";
        evt.currentTarget.className += " active";
    }
    
    // Accordion functionality
    document.addEventListener('DOMContentLoaded', function() {
        var acc = document.getElementsByClassName("accordion");
        var i;

        for (i = 0; i < acc.length; i++) {
            acc[i].addEventListener("click", function() {
                this.classList.toggle("active");
                var panel = this.nextElementSibling;
                if (panel.style.maxHeight) {
                    panel.style.maxHeight = null;
                } else {
                    panel.style.maxHeight = panel.scrollHeight + "px";
                }
            });
        }
    });
    </script>
</body>
</html>
"@
}

        # Save HTML to file
        $htmlTemplate | Out-File -FilePath $reportPath -Encoding utf8
        
        Write-Log -Message "HTML report generated: $reportPath" -Level "SUCCESS"
        return $reportPath
    }
    catch {
        Write-Log -Message "Failed to generate HTML report: $_" -Level "ERROR" -ErrorCode "ERR018"
        return $null
    }
}

#region Functions - Migration and Reporting

function New-MigrationBatch {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$ReadyMailboxes,
        
        [Parameter(Mandatory = $false)]
        [switch]$DryRun
    )
    
    try {
        if ($ReadyMailboxes.Count -eq 0) {
            Write-Log -Message "No mailboxes are ready for migration" -Level "WARNING" -ErrorCode "ERR017"
            return $null
        }
        
        # Log the number of mailboxes and their status
        Write-Log -Message "Preparing migration batch for $($ReadyMailboxes.Count) mailboxes..." -Level "INFO"
        
        # Check if we're using custom BadItemLimit recommendations
        $useBadItemLimits = $script:Config.UseBadItemLimitRecommendations
        if ($useBadItemLimits) {
            Write-Log -Message "Will use recommended BadItemLimit values for mailboxes based on risk assessment" -Level "INFO"
        }
        
        if ($DryRun) {
            Write-Log -Message "DRY RUN: Would create migration batch with the following settings:" -Level "INFO"
            Write-Log -Message "  Name: $script:BatchName" -Level "INFO"
            Write-Log -Message "  Migration Endpoint: $($script:Config.MigrationEndpointName)" -Level "INFO"
            Write-Log -Message "  Target Delivery Domain: $($script:Config.TargetDeliveryDomain)" -Level "INFO"
            Write-Log -Message "  Complete After: $((Get-Date).AddDays($script:Config.CompleteAfterDays).ToUniversalTime())" -Level "INFO"
            Write-Log -Message "  Notification Emails: $($script:Config.NotificationEmails -join ', ')" -Level "INFO"
            Write-Log -Message "  Mailboxes: $($ReadyMailboxes.Count)" -Level "INFO"
            
            # Print a sample of mailboxes
            $sampleSize = [Math]::Min(5, $ReadyMailboxes.Count)
            $sample = $ReadyMailboxes | Select-Object -First $sampleSize
            
            Write-Log -Message "Sample of mailboxes that would be migrated:" -Level "INFO"
            foreach ($mailbox in $sample) {
                $badItemLimit = if ($useBadItemLimits -and $mailbox.RecommendedBadItemLimit -gt 0) {
                    ", BadItemLimit: $($mailbox.RecommendedBadItemLimit)"
                } else { "" }
                
                Write-Log -Message "  - $($mailbox.DisplayName) ($($mailbox.EmailAddress))$badItemLimit" -Level "INFO"
            }
            
            if ($ReadyMailboxes.Count -gt $sampleSize) {
                Write-Log -Message "  ... and $($ReadyMailboxes.Count - $sampleSize) more" -Level "INFO"
            }
            
            # Return a mock migration batch
            return [PSCustomObject]@{
                Identity = "$script:BatchName (DRY RUN)"
                Status = "Not created - Dry Run"
                MigrationEndpoint = $script:Config.MigrationEndpointName
                TargetDeliveryDomain = $script:Config.TargetDeliveryDomain
                CompleteAfter = (Get-Date).AddDays($script:Config.CompleteAfterDays).ToUniversalTime()
                MailboxCount = $ReadyMailboxes.Count
                IsDryRun = $true
            }
        }
        
        # Create a temporary CSV file with just the email addresses
        $tempCsvPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "$([Guid]::NewGuid().ToString())-$script:BatchName-ready.csv"
        
        try {
            # Track the temp file for cleanup
            $script:ResourcesCreated += @{
                Type = "TempFile"
                Path = $tempCsvPath
                Created = Get-Date
            }
            
            # If using BadItemLimit recommendations, we need to create move requests individually
            # Otherwise, we can use the standard CSV approach
            if ($useBadItemLimits) {
                Write-Log -Message "Creating individual move requests with customized BadItemLimit values..." -Level "INFO"
                
                # Get migration endpoint
                $migrationEndpoint = Get-MigrationEndpoint -Identity $script:Config.MigrationEndpointName
                
                # Create a new batch (empty) first
                $completeAfterDate = (Get-Date).AddDays($script:Config.CompleteAfterDays).ToUniversalTime()
                
                $newBatchParams = @{
                    Name = $script:BatchName
                    SourceEndpoint = $migrationEndpoint.Identity
                    TargetDeliveryDomain = $script:Config.TargetDeliveryDomain
                    CompleteAfter = $completeAfterDate
                    StartAfter = (Get-Date).AddMinutes($script:Config.StartAfterMinutes)
                    NotificationEmails = $script:Config.NotificationEmails
                    AutoStart = $false  # Don't auto-start since we'll be adding mailboxes manually
                }
                
                $newBatch = New-MigrationBatch @newBatchParams
                Write-Log -Message "Created base migration batch '$script:BatchName'" -Level "SUCCESS"
                
                # Now add mailboxes to the batch with individual BadItemLimit values
                $successCount = 0
                $errorCount = 0
                
                foreach ($mailbox in $ReadyMailboxes) {
                    try {
                        $badItemLimit = if ($mailbox.RecommendedBadItemLimit -gt 0) {
                            $mailbox.RecommendedBadItemLimit
                        } else {
                            10  # Default value
                        }
                        
                        $moveRequestParams = @{
                            Identity = $mailbox.EmailAddress
                            BatchName = $script:BatchName
                            TargetDeliveryDomain = $script:Config.TargetDeliveryDomain
                            BadItemLimit = $badItemLimit
                        }
                        
                        New-MoveRequest @moveRequestParams | Out-Null
                        $successCount++
                        
                        Write-Log -Message "Added mailbox $($mailbox.EmailAddress) to batch with BadItemLimit $badItemLimit" -Level "DEBUG"
                    }
                    catch {
                        $errorCount++
                        Write-Log -Message "Failed to add mailbox $($mailbox.EmailAddress) to batch: $_" -Level "ERROR"
                    }
                }
                
                Write-Log -Message "Added $successCount of $($ReadyMailboxes.Count) mailboxes to batch '$script:BatchName' (Errors: $errorCount)" -Level "INFO"
                
                # Start the batch if all went well
                if ($successCount -gt 0) {
                    Start-MigrationBatch -Identity $script:BatchName
                    Write-Log -Message "Started migration batch '$script:BatchName'" -Level "SUCCESS"
                }
                else {
                    Write-Log -Message "No mailboxes were successfully added to batch '$script:BatchName'" -Level "ERROR" -ErrorCode "ERR010"
                }
                
                # Wait for batch to initialize
                $timeoutMinutes = $script:Config.BatchCreationTimeoutMinutes
                $endTime = (Get-Date).AddMinutes($timeoutMinutes)
                
                Write-Log -Message "Waiting for migration batch to initialize (timeout: $timeoutMinutes minutes)..." -Level "INFO"
                
                while ((Get-Date) -lt $endTime) {
                    $batchStatus = Get-MigrationBatch -Identity $script:BatchName -ErrorAction SilentlyContinue
                    
                    if ($batchStatus -and $batchStatus.Status -ne "Created") {
                        $newBatch = $batchStatus
                        break
                    }
                    
                    Write-Log -Message "Waiting for migration batch initialization..." -Level "DEBUG"
                    Start-Sleep -Seconds 5
                }
            }
            else {
                # Standard approach - create CSV and use it for batch creation
                
                # Create the CSV with explicit encoding
                "EmailAddress" | Out-File -FilePath $tempCsvPath -Encoding utf8
                $ReadyMailboxes | ForEach-Object { $_.EmailAddress } | Out-File -FilePath $tempCsvPath -Append -Encoding utf8
                
                Write-Log -Message "Created temporary CSV for migration batch at: $tempCsvPath" -Level "DEBUG"
                
                # Get migration endpoint
                $migrationEndpoint = Get-MigrationEndpoint -Identity $script:Config.MigrationEndpointName
                
                # Set complete after date
                $completeAfterDate = (Get-Date).AddDays($script:Config.CompleteAfterDays).ToUniversalTime()
                
                # Create the migration batch
                Write-Log -Message "Creating migration batch '$script:BatchName'..." -Level "INFO"
                
                $newBatchParams = @{
                    Name = $script:BatchName
                    SourceEndpoint = $migrationEndpoint.Identity
                    TargetDeliveryDomain = $script:Config.TargetDeliveryDomain
                    CSVData = [System.IO.File]::ReadAllBytes($tempCsvPath)
                    CompleteAfter = $completeAfterDate
                    StartAfter = (Get-Date).AddMinutes($script:Config.StartAfterMinutes)
                    NotificationEmails = $script:Config.NotificationEmails
                    AutoStart = $true
                }

                $newBatch = New-MigrationBatch @newBatchParams
                
                if ($newBatch -and -not $newBatchParams.AutoStart) {
                    Write-Log -Message "Starting migration batch '$script:BatchName'..." -Level "INFO"
                    Start-MigrationBatch -Identity $script:BatchName
                    Write-Log -Message "Migration batch '$script:BatchName' started successfully" -Level "SUCCESS"
                }
                
                # Wait for batch to be created and retrieve status
                $timeoutMinutes = $script:Config.BatchCreationTimeoutMinutes
                $endTime = (Get-Date).AddMinutes($timeoutMinutes)
                $batchCreated = $false
                
                Write-Log -Message "Waiting for migration batch to be created (timeout: $timeoutMinutes minutes)..." -Level "INFO"
                
                while ((Get-Date) -lt $endTime) {
                    $batchStatus = Get-MigrationBatch -Identity $script:BatchName -ErrorAction SilentlyContinue
                    
                    if ($batchStatus) {
                        $batchCreated = $true
                        $newBatch = $batchStatus
                        break
                    }
                    
                    Write-Log -Message "Waiting for migration batch creation..." -Level "DEBUG"
                    Start-Sleep -Seconds 5
                }
                
                if (-not $batchCreated) {
                    Write-Log -Message "Migration batch creation timed out after $timeoutMinutes minutes" -Level "WARNING"
                    Write-Log -Message "Check the Exchange Admin Center to verify if the batch was created" -Level "WARNING"
                }
                else {
                    Write-Log -Message "Migration batch '$script:BatchName' created successfully with status: $($newBatch.Status)" -Level "SUCCESS"
                    Write-Log -Message "Batch will start automatically in $($script:Config.StartAfterMinutes) minutes" -Level "INFO"
                    Write-Log -Message "Batch will complete after: $completeAfterDate" -Level "INFO"
                }
            }
            
            return $newBatch
        }
        finally {
            # Note: We don't remove the temp file here, it will be handled by the cleanup function
            # This ensures it's removed even if there's an error
        }
    }
    catch {
        Write-Log -Message "Failed to create migration batch: $_" -Level "ERROR" -ErrorCode "ERR010"
        Write-Log -Message "Troubleshooting: Verify permissions and migration endpoint configuration" -Level "ERROR"
        return $null
    }
}

function Invoke-ParallelMailboxValidation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$Mailboxes,
        
        [Parameter(Mandatory = $false)]
        [int]$ThrottleLimit = 5,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 2,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 3
    )
    
    try {
        Write-Log -Message "Starting parallel mailbox validation for $($Mailboxes.Count) mailboxes with throttle limit of $ThrottleLimit" -Level "INFO"
        
        # Create a progress counter
        $script:progressCount = 0
        $totalCount = $Mailboxes.Count
        
        # Create a results array
        $results = New-Object System.Collections.ArrayList
        
        # Use PowerShell 7+ parallel processing if available
        if ($PSVersionTable.PSVersion.Major -ge 7) {
            Write-Log -Message "Using PowerShell 7+ parallel processing" -Level "INFO"
            
            # Create a script for parallel execution
            $parallelScript = {
                param($EmailAddress, $RetryAttempts, $RetryDelay, $RunspaceId)
                
                # Function to handle retries
                function Invoke-WithRetry {
                    param(
                        [scriptblock]$ScriptBlock,
                        [int]$MaxRetries,
                        [int]$DelaySeconds
                    )
                    
                    $retryCount = 0
                    $success = $false
                    $result = $null
                    
                    while (-not $success -and $retryCount -le $MaxRetries) {
                        try {
                            if ($retryCount -gt 0) {
                                Write-Host "Retry $retryCount for operation..."
                                Start-Sleep -Seconds $DelaySeconds
                            }
                            
                            $result = & $ScriptBlock
                            $success = $true
                        }
                        catch {
                            $retryCount++
                            if ($retryCount -gt $MaxRetries) {
                                throw $_
                            }
                        }
                    }
                    
                    return $result
                }
                
                try {
                    # Execute the test with retry logic
                    $result = Invoke-WithRetry -ScriptBlock {
                        Test-MailboxMigrationReadiness -EmailAddress $EmailAddress
                    } -MaxRetries $RetryAttempts -DelaySeconds $RetryDelay
                    
                    return $result
                }
                catch {
                    # Create a failure result with enhanced error details
                    $result = [PSCustomObject]@{
                        EmailAddress = $EmailAddress
                        DisplayName = $EmailAddress.Split('@')[0]
                        Errors = @("Failed in runspace ID: $RunspaceId - Error: $($_.Exception.Message)")
                        ErrorCodes = @("ERR999")
                        Warnings = @()
                        OverallStatus = "Failed"
                        ExceptionDetails = $_.Exception.ToString() # Include full exception details
                    }
                    
                    return $result
                }
            }
            
            # Process mailboxes in parallel using a runspace pool
            $runspacePool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
            $runspacePool.Open()
            
            $runspaces = @()
            
            # Create and start runspaces for each mailbox
            foreach ($mailbox in $Mailboxes) {
                $powerShell = [powershell]::Create().AddScript($parallelScript)
                $null = $powerShell.AddParameter("EmailAddress", $mailbox.EmailAddress)
                $null = $powerShell.AddParameter("RetryAttempts", $RetryCount)
                $null = $powerShell.AddParameter("RetryDelay", $RetryDelaySeconds)
                $null = $powerShell.AddParameter("RunspaceId", [guid]::NewGuid().ToString()) # Add unique ID for each runspace
                $powerShell.RunspacePool = $runspacePool
                
                $runspaces += [PSCustomObject]@{
                    PowerShell = $powerShell
                    Runspace = $powerShell.BeginInvoke()
                    EmailAddress = $mailbox.EmailAddress
                }
            }
            
            # Process results as they complete
            $completed = 0
            do {
                foreach ($runspace in $runspaces.Where({ $_.Runspace.IsCompleted -eq $true -and $_.Processed -ne $true })) {
                    $result = $runspace.PowerShell.EndInvoke($runspace.Runspace)
                    $runspace.PowerShell.Dispose()
                    $runspace.Processed = $true
                    $completed++
                    
                    if ($null -ne $result) {
                        $null = $results.Add($result)
                        
                        # Output status
                        $statusColor = switch ($result.OverallStatus) {
                            "Ready" { "Green" }
                            "Warning" { "Yellow" }
                            "Failed" { "Red" }
                            default { "White" }
                        }
                        
                        Write-Host "  - $($result.DisplayName) ($($result.EmailAddress)): " -NoNewline
                        Write-Host "$($result.OverallStatus)" -ForegroundColor $statusColor
                        
                        # Log detailed error info if available
                        if ($result.ExceptionDetails) {
                            Write-Log -Message "Detailed error for $($result.EmailAddress): $($result.ExceptionDetails)" -Level "DEBUG"
                        }
                    }
                    
                    # Update progress
                    Write-Progress -Activity "Testing mailbox migration readiness" `
                        -Status "Processed $completed of $($Mailboxes.Count)" `
                        -PercentComplete (($completed / $Mailboxes.Count) * 100)
                }
                
                if ($completed -lt $Mailboxes.Count) {
                    Start-Sleep -Milliseconds 100
                }
            } while ($completed -lt $Mailboxes.Count)
            
            # Clean up
            Write-Progress -Activity "Testing mailbox migration readiness" -Completed
            $runspacePool.Close()
            $runspacePool.Dispose()
        }
        else {
            # Sequential processing for older PowerShell versions
            # Rest of the existing function unchanged...
        }
        
        Write-Log -Message "Completed validation of $($results.Count) mailboxes" -Level "INFO"
        return $results
    }
    catch {
        Write-Log -Message "Failed to perform parallel mailbox validation: $_" -Level "ERROR"
        return $null
    }
}
#endregion Functions - Migration and Reporting

#region Main Script
# Initialize Additional Validation
Write-Host "Initializing additional mailbox validation checks..." -ForegroundColor Cyan
Write-Log -Message "Additional validation functions loaded" -Level "INFO"

# Script banner
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "  Exchange Online Migration Script" -ForegroundColor Cyan
Write-Host "  Version: $script:ScriptVersion" -ForegroundColor Cyan
Write-Host "  Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""

# Check PowerShell version and recommend PowerShell 7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "╔════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║ RECOMMENDATION: This script will work better with PowerShell 7+             ║" -ForegroundColor Yellow
    Write-Host "║                                                                            ║" -ForegroundColor Yellow
    Write-Host "║ Current PowerShell Version: $($PSVersionTable.PSVersion.ToString().PadRight(41)) ║" -ForegroundColor Yellow
    Write-Host "║                                                                            ║" -ForegroundColor Yellow
    Write-Host "║ Benefits of PowerShell 7+:                                                 ║" -ForegroundColor Yellow
    Write-Host "║ - Significantly faster parallel processing                                 ║" -ForegroundColor Yellow
    Write-Host "║ - Better error handling                                                    ║" -ForegroundColor Yellow
    Write-Host "║ - Improved performance for large mailbox migrations                        ║" -ForegroundColor Yellow
    Write-Host "╚════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
    
    Write-Log -Message "PowerShell 7+ is recommended for optimal performance. Current version: $($PSVersionTable.PSVersion)" -Level "WARNING"
}

# Check for required dependencies
$dependenciesOK = Test-Dependencies
if (-not $dependenciesOK) {
    Write-Host "Required dependencies are missing. Please install them and try again." -ForegroundColor Red
    exit 1
}

# Load configuration
$script:Config = Import-MigrationConfig -ConfigPath $ConfigPath
if (-not $script:Config) {
    Write-Host "Failed to load configuration. Please check the configuration file and try again." -ForegroundColor Red
    exit 1
}

# Initialize environment
$envInitialized = Initialize-MigrationEnvironment
if (-not $envInitialized) {
    Write-Host "Failed to initialize environment. Please check the log for details." -ForegroundColor Red
    exit 1
}

if ($envInitialized) {
    Remove-OldLogs
}

# Display script mode
if ($DryRun) {
    Write-Host "Running in DRY RUN mode - no changes will be made" -ForegroundColor Yellow
    Write-Log -Message "Running in DRY RUN mode - validation only, no migration batch will be created" -Level "WARNING"
}

# Connect to Exchange Online and Microsoft Graph
$connected = Connect-MigrationServices
if (-not $connected) {
    Write-Host "Failed to connect to migration services. Please check the log for details." -ForegroundColor Red
    exit 1
}

# Import migration batch CSV
try {
    $migrationBatch = Import-Csv -Path $BatchFilePath -ErrorAction Stop
    $mailboxCount = $migrationBatch.Count
    Write-Log -Message "Imported $mailboxCount mailboxes from CSV" -Level "INFO"
}
catch {
    Write-Log -Message "Failed to import migration batch CSV: $_" -Level "ERROR" -ErrorCode "ERR007"
    Write-Host "Failed to import migration batch CSV. Please check the file and try again." -ForegroundColor Red
    exit 1
}

# Validate mailboxes
Write-Log -Message "Starting mailbox validation..." -Level "INFO"
$testResults = Invoke-ParallelMailboxValidation -Mailboxes $migrationBatch -ThrottleLimit $MaxConcurrentMailboxes -RetryCount 2 -RetryDelaySeconds 3

# Generate summary
$readyCount = ($testResults | Where-Object { $_.OverallStatus -eq "Ready" }).Count
$warningCount = ($testResults | Where-Object { $_.OverallStatus -eq "Warning" }).Count
$failedCount = ($testResults | Where-Object { $_.OverallStatus -eq "Failed" }).Count

Write-Log -Message "Mailbox readiness summary:" -Level "INFO"
Write-Log -Message "  Ready: $readyCount" -Level "INFO"
Write-Log -Message "  Warning: $warningCount" -Level "INFO"
Write-Log -Message "  Failed: $failedCount" -Level "INFO"

# Generate HTML report
$reportPath = Export-HTMLReport -TestResults $testResults
if ($reportPath) {
    Write-Host "HTML report generated: $reportPath" -ForegroundColor Green
    
    # Try to open the report in the default browser
    try {
        Start-Process $reportPath
    }
    catch {
        Write-Log -Message "Could not open report automatically: $_" -Level "WARNING"
    }
}

# If running in dry run mode, skip migration batch creation
if ($DryRun) {
    $mockBatch = [PSCustomObject]@{
        Identity = "$script:BatchName (DRY RUN)"
        Status = "Not created - Dry Run"
        IsDryRun = $true
    }
    
    $reportPath = Export-HTMLReport -TestResults $testResults -MigrationBatch $mockBatch
    
    Write-Host ""
    Write-Host "Dry run completed. No migration batch was created." -ForegroundColor Yellow
    Write-Host "Check the HTML report for detailed results: $reportPath" -ForegroundColor Yellow
    Write-Log -Message "Dry run completed successfully. HTML report: $reportPath" -Level "SUCCESS"
    
    exit 0
}

# Ask user if they want to proceed with migration if not forced
if (-not $Force) {
    $title = "Start Migration Batch"
    $message = @"
Migration readiness check completed:
- Ready: $readyCount
- Warning: $warningCount
- Failed: $failedCount

Do you want to proceed with the migration batch creation?
- 'Yes' will create a migration batch for all ready mailboxes (status = 'Ready')
- 'Force' will create a migration batch including mailboxes with warnings
- 'No' will exit without creating a migration batch
"@

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Create migration batch for ready mailboxes"
    $force = New-Object System.Management.Automation.Host.ChoiceDescription "&Force", "Create migration batch including mailboxes with warnings"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not create migration batch"

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $force, $no)
    $result = $host.ui.PromptForChoice($title, $message, $options, 2)
}
else {
    # If -Force parameter is used, automatically choose the "Force" option
    $result = 1
    Write-Log -Message "Force parameter specified, automatically including mailboxes with warnings" -Level "WARNING"
}

switch ($result) {
    0 {
        # Yes - Create migration batch for ready mailboxes
        $readyMailboxes = $testResults | Where-Object { $_.OverallStatus -eq "Ready" }
        
        if ($readyMailboxes.Count -eq 0) {
            Write-Log -Message "No mailboxes are ready for migration. Exiting without creating a batch." -Level "WARNING" -ErrorCode "ERR017"
            Write-Host "No mailboxes are ready for migration. Check the HTML report for details." -ForegroundColor Red
        }
        else {
            $migrationBatch = New-MigrationBatch -ReadyMailboxes $readyMailboxes
            
            if ($migrationBatch) {
                # Generate updated report with batch ID
                $reportPath = Export-HTMLReport -TestResults $testResults -MigrationBatch $migrationBatch
                Write-Host "Updated HTML report generated: $reportPath" -ForegroundColor Green
                
                # Try to open the updated report
                try {
                    Start-Process $reportPath
                }
                catch {
                    Write-Log -Message "Could not open updated report automatically: $_" -Level "WARNING"
                }
            }
        }
    }
    1 {
        # Force - Create migration batch including mailboxes with warnings
        $readyAndWarningMailboxes = $testResults | Where-Object { $_.OverallStatus -eq "Ready" -or $_.OverallStatus -eq "Warning" }
        
        if ($readyAndWarningMailboxes.Count -eq 0) {
            Write-Log -Message "No mailboxes are eligible for migration (even with warnings). Exiting without creating a batch." -Level "WARNING" -ErrorCode "ERR017"
            Write-Host "No mailboxes are eligible for migration. Check the HTML report for details." -ForegroundColor Red
        }
        else {
            $migrationBatch = New-MigrationBatch -ReadyMailboxes $readyAndWarningMailboxes
            
            if ($migrationBatch) {
                # Generate updated report with batch ID
                $reportPath = Export-HTMLReport -TestResults $testResults -MigrationBatch $migrationBatch
                Write-Host "Updated HTML report generated: $reportPath" -ForegroundColor Green
                
                # Try to open the updated report
                try {
                    Start-Process $reportPath
                }
                catch {
                    Write-Log -Message "Could not open updated report automatically: $_" -Level "WARNING"
                }
            }
        }
    }
    2 {
        # No - Do not create migration batch
        Write-Log -Message "User chose not to create a migration batch. Exiting script." -Level "INFO"
        Write-Host "Migration batch creation cancelled." -ForegroundColor Yellow
    }
}

Write-Log -Message "Exchange Online Migration Script completed" -Level "SUCCESS"
Write-Host ""
Write-Host "Script completed. Log file: $script:LogFile" -ForegroundColor Green
#endregion Main Script