function Test-TokenExpiration {
    <#
    .SYNOPSIS
        Checks if authentication tokens are about to expire and refreshes them if needed.
    
    .DESCRIPTION
        Monitors the age of authentication tokens for Exchange Online and Microsoft Graph
        connections, and forces a refresh when they approach expiration to prevent
        authentication errors during long-running operations.
    
    .PARAMETER ForceRefresh
        When specified, forces a token refresh regardless of token age.
    
    .PARAMETER TokenLifetimeMinutes
        The expected lifetime of authentication tokens in minutes. Default is 50 minutes.
    
    .PARAMETER RefreshBufferMinutes
        Number of minutes before expiration to trigger refresh. Default is 5 minutes.
    
    .EXAMPLE
        Test-TokenExpiration
    
    .EXAMPLE
        Test-TokenExpiration -ForceRefresh
    
    .OUTPUTS
        [bool] Returns $true if tokens are valid or successfully refreshed, $false otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ForceRefresh,
        
        [Parameter(Mandatory = $false)]
        [int]$TokenLifetimeMinutes = 0,
        
        [Parameter(Mandatory = $false)]
        [int]$RefreshBufferMinutes = 0
    )
    
    try {
        # Use config values if not specified in parameters
        if ($TokenLifetimeMinutes -eq 0 -and $script:Config.TokenLifetimeMinutes) {
            $TokenLifetimeMinutes = $script:Config.TokenLifetimeMinutes
        }
        elseif ($TokenLifetimeMinutes -eq 0) {
            # Default value if not in config
            $TokenLifetimeMinutes = 50  # Typical token lifetime is 60 minutes
        }
        
        if ($RefreshBufferMinutes -eq 0 -and $script:Config.TokenRefreshBuffer) {
            $RefreshBufferMinutes = $script:Config.TokenRefreshBuffer
        }
        elseif ($RefreshBufferMinutes -eq 0) {
            # Default value if not in config
            $RefreshBufferMinutes = 5
        }
        
        # Check if token might be expiring soon based on the last connection time
        if ($ForceRefresh -or 
            $global:LastConnectionTime -eq $null -or 
            ((Get-Date) - $global:LastConnectionTime).TotalMinutes -gt ($TokenLifetimeMinutes - $RefreshBufferMinutes)) {
            
            $minutesRemaining = if ($global:LastConnectionTime) {
                $TokenLifetimeMinutes - ((Get-Date) - $global:LastConnectionTime).TotalMinutes
            } else {
                "unknown"
            }
            
            if ($ForceRefresh) {
                Write-Log -Message "Token refresh requested - reconnecting services" -Level "INFO"
            }
            else {
                Write-Log -Message "Authentication token approaching expiration (approximately $minutesRemaining minutes remaining) - reconnecting services" -Level "INFO"
            }
            
            # Force reconnect to refresh tokens
            return Connect-MigrationServices -ForceReconnect
        }
        
        # Check Exchange Online token by trying a lightweight command
        try {
            $null = Get-AcceptedDomain -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -like "*token*" -or 
                $_.Exception.Message -like "*unauthorized*" -or 
                $_.Exception.Message -like "*authentication*" -or
                $_.Exception.Message -like "*401*" -or 
                $_.Exception.Message -like "*session*expired*") {
                
                Write-Log -Message "Exchange Online token has expired - reconnecting services" -Level "WARNING"
                return Connect-MigrationServices -ForceReconnect
            }
            else {
                # This is a non-auth related error, don't attempt reconnection
                Write-Log -Message "Error checking Exchange Online connection: $_" -Level "ERROR"
                return $false
            }
        }
        
        # Check Microsoft Graph token by trying a lightweight command
        try {
            $null = Get-MgUser -Top 1 -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -like "*token*" -or 
                $_.Exception.Message -like "*unauthorized*" -or 
                $_.Exception.Message -like "*authentication*" -or
                $_.Exception.Message -like "*401*" -or 
                $_.Exception.Message -like "*session*expired*") {
                
                Write-Log -Message "Microsoft Graph token has expired - reconnecting services" -Level "WARNING"
                return Connect-MigrationServices -ForceReconnect
            }
            else {
                # This is a non-auth related error, don't attempt reconnection
                Write-Log -Message "Error checking Microsoft Graph connection: $_" -Level "ERROR"
                return $false
            }
        }
        
        # If we get here, both tokens are valid
        return $true
    }
    catch {
        Write-Log -Message "Error checking token expiration: $_" -Level "ERROR"
        return $false
    }
}
