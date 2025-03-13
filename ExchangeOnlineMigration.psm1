# ExchangeOnlineMigration.psm1
# Module loader for Exchange Online Migration module

# Establish script global variables
$script:ScriptVersion = "3.0.0"
$script:LogFile = $null
$script:Config = $null
$script:LogBuffer = @()

# Get the module path
$modulePath = $PSScriptRoot

# Load private functions first (recursively from all subdirectories)
$privatePath = Join-Path -Path $modulePath -ChildPath "Private"
if (Test-Path -Path $privatePath) {
    $privateFiles = Get-ChildItem -Path $privatePath -Filter "*.ps1" -Recurse
    foreach ($file in $privateFiles) {
        try {
            . $file.FullName
            Write-Verbose "Imported private file: $($file.FullName)"
        }
        catch {
            Write-Error "Failed to import private file $($file.FullName): $_"
        }
    }
}

# Load public functions
$publicPath = Join-Path -Path $modulePath -ChildPath "Public"
$publicFunctions = @()

if (Test-Path -Path $publicPath) {
    $publicFiles = Get-ChildItem -Path $publicPath -Filter "*.ps1" -Recurse
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

    # New error codes for advanced validation
    "ERR030" = "AD sync verification failed - mailbox not properly synchronized"
    "ERR031" = "User not found in Microsoft Graph API"
    "ERR032" = "Soft-deleted mailbox found in Exchange Online"
    "ERR033" = "Mail user with matching identity exists in Exchange Online"
    "ERR034" = "Mailbox is configured as a journal recipient"
    "ERR035" = "Mailbox has email addresses in unverified domains"
    "ERR036" = "Mailbox is in a soft-deleted state"
    "ERR037" = "Found soft-deleted mailboxes that could conflict with migration"
    "ERR038" = "Found alias or email address conflicts that may prevent migration"
    "ERR039" = "Archive mailbox size exceeds license limit"
    "ERR040" = "Namespace conflict detected"
    "ERR041" = "Problematic folder names detected"
    "ERR042" = "Recoverable items folder size exceeds recommended limit"
    "ERR043" = "Very old items detected in mailbox"
    "ERR044" = "High mailbox activity level detected"
    "ERR045" = "Complex folder permission structure detected"
    
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