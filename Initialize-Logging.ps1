function Initialize-Logging {
    <#
    .SYNOPSIS
        Initializes the logging system for the migration process.
    
    .DESCRIPTION
        Sets up the logging environment including log file creation and flushing
        buffered log messages. Creates log directories if they don't exist.
    
    .PARAMETER LogPath
        Path to the log file. If directory doesn't exist, it will be created.
    
    .PARAMETER FlushBuffer
        When specified, flushes any buffered log messages to the new log file.
    
    .PARAMETER LogLevel
        The minimum log level to record. Default is "INFO".
    
    .EXAMPLE
        Initialize-Logging -LogPath "C:\Migration\Logs\migration.log"
    
    .EXAMPLE
        Initialize-Logging -LogPath "C:\Migration\Logs\migration.log" -LogLevel "DEBUG"
    
    .OUTPUTS
        [bool] Returns $true if logging was initialized successfully, $false otherwise.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$FlushBuffer,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL")]
        [string]$LogLevel = "INFO"
    )
    
    try {
        # Create log directory if it doesn't exist
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            # We can't use Write-Log yet since it's not initialized
            Write-Host "Created log directory: $logDir" -ForegroundColor Green
        }
        
        # Create or append header to log file
        $logHeader = @"
======================================================================
Exchange Online Migration - Log Started: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Script Version: $($script:ScriptVersion)
PowerShell Version: $($PSVersionTable.PSVersion)
Computer Name: $env:COMPUTERNAME
User: $env:USERNAME
Log Level: $LogLevel
======================================================================

"@
        
        # Set the global log file path
        $script:LogFile = $LogPath
        
        # Create the log file with header
        $logHeader | Out-File -FilePath $LogPath -Encoding utf8
        
        # Set the minimum log level
        $script:LogLevel = $LogLevel
        
        # Flush any buffered log messages if requested
        if ($FlushBuffer -and $script:LogBuffer) {
            $script:LogBuffer | Out-File -FilePath $LogPath -Append -Encoding utf8
            $script:LogBuffer = @()
            Write-Host "Flushed buffered log messages to log file" -ForegroundColor Green
        }
        
        Write-Log -Message "Logging initialized at $LogPath with log level $LogLevel" -Level "INFO"
        return $true
    }
    catch {
        # Can't use Write-Log here if it failed
        Write-Host "Failed to initialize logging: $_" -ForegroundColor Red
        return $false
    }
}
