function Invoke-WithTokenRefresh {
    <#
    .SYNOPSIS
        Executes a script block with automatic token refresh on authentication failures.
    
    .DESCRIPTION
        Attempts to execute the script block and automatically refreshes the authentication
        token and retries if an authentication error occurs. This helps avoid token expiration
        issues during long-running operations.
    
    .PARAMETER ScriptBlock
        The script block to execute.
    
    .PARAMETER MaxRetries
        Maximum number of retry attempts after token refresh. Default is 1.
    
    .PARAMETER ErrorPatterns
        Array of string patterns to identify authentication-related errors. Defaults to common patterns.
    
    .EXAMPLE
        Invoke-WithTokenRefresh -ScriptBlock { Get-Mailbox -Identity "user@contoso.com" }
    
    .EXAMPLE
        Invoke-WithTokenRefresh -ScriptBlock { Get-MigrationBatch } -MaxRetries 2
    
    .OUTPUTS
        Returns the output of the script block if successful.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 1,
        
        [Parameter(Mandatory = $false)]
        [string[]]$ErrorPatterns = @(
            "token",
            "unauthorized",
            "authentication",
            "401",
            "session expired"
        )
    )
    
    try {
        Write-Verbose "Executing script block with token refresh support"
        
        # First attempt - try to execute without refresh
        return & $ScriptBlock
    }
    catch {
        $errorMessage = $_.Exception.Message
        $retryCount = 0
        $isAuthError = $false
        
        # Check if this is an auth-related error
        foreach ($pattern in $ErrorPatterns) {
            if ($errorMessage -like "*$pattern*") {
                $isAuthError = $true
                break
            }
        }
        
        if (-not $isAuthError) {
            # Not an auth error, just rethrow
            Write-Verbose "Error is not authentication-related. Rethrowing original exception."
            throw $_
        }
        
        # Auth error detected, attempt to refresh and retry
        while ($retryCount -lt $MaxRetries) {
            $retryCount++
            
            try {
                Write-Log -Message "Authentication error detected, refreshing token and retrying (Attempt $retryCount of $MaxRetries)..." -Level "WARNING"
                
                # Call token refresh function
                $refreshResult = Test-TokenExpiration -ForceRefresh
                
                if (-not $refreshResult) {
                    Write-Log -Message "Token refresh failed, aborting operation" -Level "ERROR"
                    throw "Failed to refresh authentication token after authentication error: $errorMessage"
                }
                
                Write-Log -Message "Token refreshed successfully, retrying operation" -Level "INFO"
                
                # Try the operation again
                return & $ScriptBlock
            }
            catch {
                if ($retryCount -ge $MaxRetries) {
                    Write-Log -Message "Maximum retries reached after authentication errors" -Level "ERROR"
                    throw "Operation failed after $MaxRetries retry attempts: $_"
                }
                
                # If this is another auth error, we'll retry again
                $stillAuthError = $false
                $newErrorMessage = $_.Exception.Message
                
                foreach ($pattern in $ErrorPatterns) {
                    if ($newErrorMessage -like "*$pattern*") {
                        $stillAuthError = $true
                        break
                    }
                }
                
                if (-not $stillAuthError) {
                    # If it's no longer an auth error, just rethrow the new error
                    throw $_
                }
                
                # It's still an auth error, we'll retry in the next loop iteration
                Write-Log -Message "Authentication error persists after token refresh. Will retry again." -Level "WARNING"
            }
        }
        
        # Should not reach here due to throws above, but just in case
        throw "Operation failed due to persistent authentication errors after $MaxRetries retries"
    }
}
