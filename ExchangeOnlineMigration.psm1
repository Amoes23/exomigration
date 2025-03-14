# ExchangeOnlineMigration.psm1
# Module loader for Exchange Online Migration module

# Establish script global variables
$script:ScriptVersion = "3.1.0"
$script:LogFile = $null
$script:Config = $null
$script:LogBuffer = @()
$script:LastConnectionTime = $null

# Get the module path
$modulePath = $PSScriptRoot

# Define the loading order for core private functions
$coreDependencyOrder = @(
    'Write-Log.ps1',
    'Import-MigrationConfig.ps1',
    'Test-DiskSpace.ps1',
    'New-MailboxTestResult.ps1',
    'Invoke-WithRetry.ps1',
    'Connect-MigrationEnvironment.ps1',
    'Test-TokenExpiration.ps1'
)

# First load the core dependencies in specified order
$corePath = Join-Path -Path $modulePath -ChildPath "Private\Core"
if (Test-Path -Path $corePath) {
    foreach ($coreFile in $coreDependencyOrder) {
        $coreFilePath = Join-Path -Path $corePath -ChildPath $coreFile
        if (Test-Path -Path $coreFilePath) {
            try {
                . $coreFilePath
                Write-Verbose "Imported core file: $coreFilePath"
            }
            catch {
                Write-Error "Failed to import core file $coreFilePath: $_"
            }
        }
        else {
            Write-Warning "Core dependency file not found: $coreFilePath"
        }
    }
    
    # Load remaining core files
    $remainingCoreFiles = Get-ChildItem -Path $corePath -Filter "*.ps1" | 
                          Where-Object { $_.Name -notin $coreDependencyOrder }
    
    foreach ($file in $remainingCoreFiles) {
        try {
            . $file.FullName
            Write-Verbose "Imported additional core file: $($file.FullName)"
        }
        catch {
            Write-Error "Failed to import core file $($file.FullName): $_"
        }
    }
}

# Load validation functions by category
$validationPaths = @(
    (Join-Path -Path $modulePath -ChildPath "Private\Validation\Common"),
    (Join-Path -Path $modulePath -ChildPath "Private\Validation\Cloud"),
    (Join-Path -Path $modulePath -ChildPath "Private\Validation\OnPremises")
)

foreach ($path in $validationPaths) {
    if (Test-Path -Path $path) {
        $validationFiles = Get-ChildItem -Path $path -Filter "*.ps1"
        foreach ($file in $validationFiles) {
            try {
                . $file.FullName
                Write-Verbose "Imported validation file: $($file.FullName)"
            }
            catch {
                Write-Error "Failed to import validation file $($file.FullName): $_"
            }
        }
    }
}

# Load reporting functions
$reportingPath = Join-Path -Path $modulePath -ChildPath "Private\Reporting"
if (Test-Path -Path $reportingPath) {
    $reportingFiles = Get-ChildItem -Path $reportingPath -Filter "*.ps1"
    foreach ($file in $reportingFiles) {
        try {
            . $file.FullName
            Write-Verbose "Imported reporting file: $($file.FullName)"
        }
        catch {
            Write-Error "Failed to import reporting file $($file.FullName): $_"
        }
    }
}

# Load public functions last
$publicPath = Join-Path -Path $modulePath -ChildPath "Public"
$publicFunctions = @()

if (Test-Path -Path $publicPath) {
    $publicFiles = Get-ChildItem -Path $publicPath -Filter "*.ps1"
    foreach ($file in $publicFiles) {
        try {
            . $file.FullName
            $publicFunctions += $file.BaseName
            Write-Verbose "Imported public function: $($file.BaseName)"
        }
        catch {
            Write-Error "Failed to import public file $($file.FullName): $_"
        }
    }
}

# Export error codes variable
$global:EXOMigrationErrorCodes = @{
    # Connection and initialization errors
    "ERR001" = "Failed to load required PowerShell modules"
    "ERR002" = "Failed to load configuration from file"
    "ERR003" = "Failed to initialize environment"
    "ERR004" = "Failed to connect to Exchange Online"
    "ERR005" = "Failed to connect to Microsoft Graph"
    "ERR006" = "Failed to validate migration endpoint"
    "ERR007" = "Failed to connect to on-premises Exchange"
    
    # Batch and file errors
    "ERR010" = "Failed to import migration batch CSV"
    "ERR011" = "Batch CSV doesn't contain required 'EmailAddress' column"
    "ERR012" = "Migration batch with same name already exists"
    "ERR013" = "Failed to create migration batch"
    
    # Mailbox configuration errors
    "ERR020" = "Invalid mailbox format"
    "ERR021" = "Missing required license"
    "ERR022" = "License provisioning error"
    "ERR023" = "Missing onmicrosoft.com email address"
    "ERR024" = "Domain not verified in tenant"
    "ERR025" = "Pending move request detected"
    "ERR026" = "No mailboxes ready for migration"
    "ERR027" = "Failed to generate HTML report"
    
    # Migration validation errors
    "ERR030" = "Mailbox contains items exceeding size limit"
    "ERR031" = "Unified Messaging is enabled and requires special handling"
    "ERR032" = "User belongs to nested groups that may require special handling"
    "ERR033" = "Mailbox has orphaned permissions"
    "ERR034" = "Mailbox has deeply nested folders that may cause issues"
    "ERR035" = "Mailbox has an excessive number of items"
    "ERR036" = "Shared calendar permissions require manual recreation"
    "ERR037" = "Special mailbox requires special handling"

    # Exchange Online specific errors
    "ERR040" = "AD sync verification failed - mailbox not properly synchronized"
    "ERR041" = "User not found in Microsoft Graph API"
    "ERR042" = "Soft-deleted mailbox found in Exchange Online"
    "ERR043" = "Mail user with matching identity exists in Exchange Online"
    "ERR044" = "Mailbox is configured as a journal recipient"
    "ERR045" = "Mailbox has email addresses in unverified domains"
    "ERR046" = "Mailbox is in a soft-deleted state"
    "ERR047" = "Found soft-deleted mailboxes that could conflict with migration"
    "ERR048" = "Found alias or email address conflicts that may prevent migration"
    "ERR049" = "Archive mailbox size exceeds license limit"
    
    # On-premises specific errors
    "ERR050" = "Legacy Exchange DN is missing or invalid"
    "ERR051" = "On-premises mailbox lacks required attributes"
    "ERR052" = "On-premises mailbox has incompatible configuration"
    
    # Common errors
    "ERR060" = "Namespace conflict detected"
    "ERR061" = "Problematic folder names detected"
    "ERR062" = "Recoverable items folder size exceeds recommended limit"
    "ERR063" = "Very old items detected in mailbox"
    "ERR064" = "High mailbox activity level detected"
    "ERR065" = "Complex folder permission structure detected"
    
    # General errors
    "ERR999" = "Unspecified error during validation"
}

# Define the functions to export
$functionsToExport = @(
    'Start-EXOMigration',
    'Test-EXOMailboxReadiness',
    'Invoke-EXOParallelValidation',
    'New-EXOMigrationReport',
    'Get-EXOMigrationStatus',
    'New-EXOMigrationBatch'
)

# Create a script block that exports everything
$exportScriptBlock = {
    # Export all the functions we want to expose
    Export-ModuleMember -Function $functionsToExport
    
    # Export aliases
    Export-ModuleMember -Alias "Start-ExchangeMigration", "Test-MailboxReadiness"
    
    # Export variables
    Export-ModuleMember -Variable "EXOMigrationErrorCodes"
}

# Execute the export script block
. $exportScriptBlock

# Write module loaded message
Write-Verbose "Exchange Online Migration Module v$script:ScriptVersion loaded"
Write-Verbose "Exported functions: $($functionsToExport -join ', ')"
