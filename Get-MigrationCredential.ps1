function Get-MigrationCredential {
    <#
    .SYNOPSIS
        Gets credentials for migration operations.
    
    .DESCRIPTION
        Retrieves credentials for Exchange Online migration operations.
        Tries to use stored credentials first, then prompts if needed.
    
    .PARAMETER CredentialName
        Name for identifying the credential in the credential store.
    
    .PARAMETER Message
        The message to display when prompting for credentials.
    
    .PARAMETER Force
        When specified, forces a new credential prompt even if credentials exist.
    
    .EXAMPLE
        $cred = Get-MigrationCredential
    
    .EXAMPLE
        $cred = Get-MigrationCredential -Force -Message "Enter admin credentials for tenant migration"
    
    .OUTPUTS
        [System.Management.Automation.PSCredential] Credential object for authentication.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$CredentialName = "ExchangeMigration",
        
        [Parameter(Mandatory = $false)]
        [string]$Message = "Enter credentials for Exchange Online Migration",
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    
    # Check if we already have the credential in this session
    if (-not $Force -and $script:MigrationCredential) {
        Write-Verbose "Using existing migration credential from session"
        return $script:MigrationCredential
    }
    
    try {
        # First try Windows Credential Manager if available
        if (-not $Force) {
            try {
                if (Get-Command -Name Get-StoredCredential -ErrorAction SilentlyContinue) {
                    $storedCred = Get-StoredCredential -Target $CredentialName -ErrorAction Stop
                    if ($storedCred) {
                        Write-Verbose "Retrieved credential from Windows Credential Manager"
                        $script:MigrationCredential = $storedCred
                        return $storedCred
                    }
                }
            }
            catch {
                Write-Verbose "Could not retrieve from credential store: $_"
            }
        }
        
        # Fall back to prompting user
        try {
            $credential = Get-Credential -Message $Message
            if ($credential) {
                $script:MigrationCredential = $credential
                
                # Try to save to credential manager if available
                try {
                    if (Get-Command -Name New-StoredCredential -ErrorAction SilentlyContinue) {
                        New-StoredCredential -Target $CredentialName -Credential $credential -Persist LocalMachine
                        Write-Verbose "Saved credential to Windows Credential Manager"
                    }
                }
                catch {
                    Write-Verbose "Could not save to credential store: $_"
                }
                
                return $credential
            }
            else {
                throw "No credential provided"
            }
        }
        catch {
            Write-Log -Message "Failed to get credential: $_" -Level "ERROR"
            throw
        }
    }
    catch {
        Write-Log -Message "Error in Get-MigrationCredential: $_" -Level "ERROR"
        throw
    }
}