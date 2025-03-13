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
    
    .PARAMETER NoTimestamp
        If set to $true, suppress the timestamp in console output. Default is $false.
    
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
        [switch]$Console = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$NoTimestamp = $false
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level]"
    
    # Add error code if specified
    if ($ErrorCode) {
        $logMessage += " [$ErrorCode]"
    }
    
    $logMessage += " $Message"
    
    # Console message can be different than file message (without timestamp if requested)
    $consoleMessage = $logMessage
    if ($NoTimestamp) {
        $consoleMessage = "[$Level]"
        if ($ErrorCode) {
            $consoleMessage += " [$ErrorCode]"
        }
        $consoleMessage += " $Message"
    }
    
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
        # Still write to log file if we have one, but skip console output
        if ($script:LogFile) {
            Add-Content -Path $script:LogFile -Value $logMessage
        }
        else {
            # If log file isn't initialized yet, buffer this message to write later
            $script:LogBuffer = $script:LogBuffer ?? @()
            $script:LogBuffer += $logMessage
        }
        return
    }
    
    # Write to console with color if requested
    if ($Console) {
        Write-Host $consoleMessage -ForegroundColor $color
    }
    
    # Append to log file if initialized
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logMessage
    }
    else {
        # If log file isn't initialized yet, buffer this message to write later
        $script:LogBuffer = $script:LogBuffer ?? @()
        $script:LogBuffer += $logMessage
    }
    
    # For critical errors, consider additional notification
    if ($Level -eq "CRITICAL" -and $script:Config -and $script:Config.NotificationEmails -and $script:Config.NotificationEmails.Count -gt 0) {
        try {
            $subject = "CRITICAL ERROR in Exchange Migration - $ErrorCode"
            $body = @"
<h2>CRITICAL ERROR in Exchange Migration</h2>
<p><strong>Error:</strong> $Message</p>
<p><strong>Error Code:</strong> $ErrorCode</p>
<p><strong>Time:</strong> $timestamp</p>
<p><strong>Script:</strong> $($MyInvocation.ScriptName)</p>
<p><strong>Line Number:</strong> $($MyInvocation.ScriptLineNumber)</p>
"@
            Send-MigrationNotification -Subject $subject -Body $body -BodyAsHtml -Priority High
        }
        catch {
            # Don't let notification failures break logging
            Write-Host "Failed to send critical error notification: $_" -ForegroundColor Red
        }
    }
}