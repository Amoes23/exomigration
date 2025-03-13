function Remove-OldLogs {
    <#
    .SYNOPSIS
        Removes old log files beyond the specified retention period.
    
    .DESCRIPTION
        Cleans up log files that are older than the specified number of days
        to prevent excessive disk space usage from log file accumulation.
    
    .PARAMETER LogPath
        Path to the log directory. Default is taken from the script configuration.
    
    .PARAMETER LogCleanupDays
        Number of days to keep logs. Default is taken from the script configuration.
    
    .PARAMETER LogFilePattern
        Pattern for matching log files. Default is "*.log".
    
    .EXAMPLE
        Remove-OldLogs
    
    .EXAMPLE
        Remove-OldLogs -LogPath "C:\Migration\Logs" -LogCleanupDays 14
    
    .OUTPUTS
        None
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [int]$LogCleanupDays = 0,
        
        [Parameter(Mandatory = $false)]
        [string]$LogFilePattern = "*.log"
    )
    
    try {
        # Use config values if parameters not provided
        if ([string]::IsNullOrEmpty($LogPath) -and $script:Config -and $script:Config.LogPath) {
            $LogPath = $script:Config.LogPath
        }
        elseif ([string]::IsNullOrEmpty($LogPath)) {
            $LogPath = ".\Logs\"
        }
        
        if ($LogCleanupDays -eq 0 -and $script:Config -and $script:Config.LogCleanupDays) {
            $LogCleanupDays = $script:Config.LogCleanupDays
        }
        elseif ($LogCleanupDays -eq 0) {
            $LogCleanupDays = 30  # Default retention period
        }
        
        # Ensure log directory exists
        if (-not (Test-Path -Path $LogPath)) {
            Write-Log -Message "Log directory does not exist: $LogPath" -Level "WARNING"
            return
        }
        
        # Calculate cutoff date
        $cutoffDate = (Get-Date).AddDays(-$LogCleanupDays)
        Write-Log -Message "Removing log files older than $LogCleanupDays days (before $cutoffDate)" -Level "INFO"
        
        # Find old log files
        $oldLogs = Get-ChildItem -Path $LogPath -Filter $LogFilePattern | 
            Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        if ($oldLogs.Count -eq 0) {
            Write-Log -Message "No log files older than $LogCleanupDays days found" -Level "INFO"
            return
        }
        
        # Remove old log files
        $removedCount = 0
        $errorCount = 0
        
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
                    $errorCount++
                    Write-Log -Message "Skipping log file $($log.Name) as it is currently in use" -Level "WARNING"
                }
            }
            catch {
                $errorCount++
                Write-Log -Message "Failed to remove log file $($log.Name): $_" -Level "WARNING"
            }
        }
        
        if ($removedCount -gt 0) {
            Write-Log -Message "Removed $removedCount out of $($oldLogs.Count) log files older than $LogCleanupDays days" -Level "SUCCESS"
        }
        
        if ($errorCount -gt 0) {
            Write-Log -Message "Failed to remove $errorCount log files" -Level "WARNING"
        }
    }
    catch {
        Write-Log -Message "Error during log cleanup: $_" -Level "ERROR"
    }
}
