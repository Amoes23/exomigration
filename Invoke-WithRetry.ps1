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
    
    .PARAMETER Activity
        Optional activity name to display in verbose logging.
    
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
        [string]$ExceptionMessage = $null,
        
        [Parameter(Mandatory = $false)]
        [string]$Activity = "Operation"
    )
    
    $retryCount = 0
    $success = $false
    $result = $null
    
    while (-not $success -and $retryCount -le $MaxRetries) {
        try {
            if ($retryCount -gt 0) {
                $sleepTime = $DelaySeconds * [Math]::Pow(2, ($retryCount - 1))  # Exponential backoff
                $sleepTime = [Math]::Min($sleepTime, 30)  # Cap at 30 seconds
                
                Write-Log -Message "Retry $retryCount/$MaxRetries for $Activity after $sleepTime second delay..." -Level "DEBUG"
                Start-Sleep -Seconds $sleepTime
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
                Write-Log -Message "$Activity failed, will retry ($retryCount/$MaxRetries): $errorMsg" -Level "WARNING"
            }
            else {
                if ($retryCount -gt $MaxRetries) {
                    Write-Log -Message "$Activity failed after $retryCount attempts: $errorMsg" -Level "ERROR"
                }
                else {
                    Write-Log -Message "$Activity failed with non-retryable exception: $errorMsg" -Level "ERROR"
                }
                throw $_
            }
        }
    }
    
    if ($success -and $retryCount -gt 0) {
        Write-Log -Message "$Activity succeeded after $retryCount retries" -Level "SUCCESS"
    }
    elseif ($success) {
        Write-Log -Message "$Activity completed successfully" -Level "DEBUG"
    }
    
    return $result
}
