function Save-MigrationState {
    <#
    .SYNOPSIS
        Saves the current migration state to a file.
    
    .DESCRIPTION
        Serializes and saves the migration state object to a JSON file,
        enabling the migration process to be resumed later if interrupted.
    
    .PARAMETER State
        The migration state object to save.
    
    .PARAMETER Path
        Path where the state file should be saved.
    
    .PARAMETER CreateBackup
        When specified, creates a backup of any existing state file before overwriting.
    
    .EXAMPLE
        Save-MigrationState -State $migrationState -Path "C:\Migration\state.json"
    
    .EXAMPLE
        Save-MigrationState -State $migrationState -Path "C:\Migration\state.json" -CreateBackup
    
    .OUTPUTS
        [bool] Returns $true if the state was saved successfully, $false otherwise.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$State,
        
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $false)]
        [switch]$CreateBackup
    )
    
    try {
        # Make sure the directory exists
        $directory = Split-Path -Path $Path -Parent
        if (-not (Test-Path -Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
            Write-Log -Message "Created directory for state file: $directory" -Level "INFO"
        }
        
        # Create backup if requested and original exists
        if ($CreateBackup -and (Test-Path -Path $Path)) {
            $backupPath = "$Path.backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
            Copy-Item -Path $Path -Destination $backupPath -Force
            Write-Log -Message "Created backup of state file: $backupPath" -Level "DEBUG"
        }
        
        # Update the last updated timestamp
        $State.LastUpdated = Get-Date
        
        # Convert to JSON with reasonable depth to capture all properties
        $json = ConvertTo-Json -InputObject $State -Depth 10
        
        # Save to file with UTF-8 encoding
        $json | Out-File -FilePath $Path -Force -Encoding utf8
        
        Write-Log -Message "Migration state saved: $($State.CurrentStage)" -Level "DEBUG"
        return $true
    }
    catch {
        Write-Log -Message "Failed to save migration state: $_" -Level "ERROR"
        return $false
    }
}
