function Connect-MigrationServices {
    <#
    .SYNOPSIS
        Connects to Exchange Online and Microsoft Graph services required for migration.
    
    .DESCRIPTION
        Establishes authenticated connections to Exchange Online Management and Microsoft Graph,
        supporting multiple authentication methods including modern auth, certificate-based auth,
        and device code flow. Verifies migration endpoint availability and tests connection health.
    
    .PARAMETER ForceReconnect
        When specified, forces a new connection even if already connected.
    
    .PARAMETER TokenLifetimeMinutes
        The expected lifetime of authentication tokens in minutes. Default is 50 minutes.
    
    .PARAMETER RetryCount
        Number of connection attempts before giving up. Default is 2.
    
    .PARAMETER Credential
        Optional PSCredential object containing authentication credentials.
    
    .EXAMPLE
        Connect-MigrationServices
    
    .EXAMPLE
        Connect-MigrationServices -ForceReconnect -Credential $cred
    
    .OUTPUTS
        [bool] True if connected successfully, False otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ForceReconnect,
        
        [Parameter(Mandatory = $false)]
        [int]$TokenLifetimeMinutes = 50,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 2,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Track if we need to reconnect based on token lifetime
        $global:LastConnectionTime = $global:LastConnectionTime ?? $null
        $tokenExpired = $false
        
        if ($global:LastConnectionTime -ne $null) {
            $timeSinceLastConnection = (Get-Date) - $global:LastConnectionTime
            $tokenExpired = $timeSinceLastConnection.TotalMinutes -gt $TokenLifetimeMinutes
            
            if ($tokenExpired) {
                Write-Log -Message "Authentication token may be expired (last connection: $global:LastConnectionTime)" -Level "WARNING"
            }
        }
        
        # Get credential if not provided
        if (-not $Credential) {
            try {
                $Credential = Get-MigrationCredential
            }
            catch {
                Write-Log -Message "Failed to get migration credentials: $_" -Level "WARNING"
                Write-Log -Message "Will attempt to connect using current context or device code authentication" -Level "INFO"
            }
        }
        
        # Connect to Exchange Online if not already connected or if force reconnect is specified
        $exchangeConnected = $false
        $retryAttempt = 0
        
        while (-not $exchangeConnected -and $retryAttempt -lt $RetryCount) {
            try {
                if (-not $ForceReconnect -and -not $tokenExpired) {
                    # Check if already connected
                    $orgConfig = Get-OrganizationConfig -ErrorAction Stop
                    Write-Log -Message "Already connected to Exchange Online organization: $($orgConfig.DisplayName)" -Level "INFO"
                    $exchangeConnected = $true
                }
                else {
                    throw "Need to establish new connection"
                }
            }
            catch {
                $retryAttempt++
                Write-Log -Message "Connecting to Exchange Online (Attempt $retryAttempt of $RetryCount)..." -Level "INFO"
                
                try {
                    # Check for existing session and remove if necessary
                    $existingSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
                    if ($existingSession) {
                        Write-Log -Message "Removing existing Exchange Online session" -Level "INFO"
                        $existingSession | Remove-PSSession -ErrorAction SilentlyContinue
                    }
                    
                    # Prepare connection parameters
                    $exoParams = @{
                        ShowBanner = $false
                        ShowProgress = $true
                        PSSessionOption = New-PSSessionOption -IdleTimeout 900000  # 15-minute timeout
                    }
                    
                    # Add credential if provided
                    if ($Credential) {
                        $exoParams.Add('Credential', $Credential)
                    }
                    
                    # Connect to Exchange Online
                    Connect-ExchangeOnline @exoParams
                    
                    $orgConfig = Get-OrganizationConfig
                    $global:LastConnectionTime = Get-Date
                    $exchangeConnected = $true
                    
                    Write-Log -Message "Connected to Exchange Online organization: $($orgConfig.DisplayName)" -Level "SUCCESS"
                }
                catch {
                    if ($retryAttempt -ge $RetryCount) {
                        Write-Log -Message "Failed to connect to Exchange Online after $RetryCount attempts: $_" -Level "ERROR" -ErrorCode "ERR004"
                        Write-Log -Message "Check your internet connection and credentials. Try manually running Connect-ExchangeOnline in a separate window." -Level "ERROR"
                    }
                    else {
                        Write-Log -Message "Connection attempt failed, retrying in 5 seconds: $_" -Level "WARNING"
                        Start-Sleep -Seconds 5
                    }
                }
            }
        }
        
        # Connect to Microsoft Graph if not already connected
        $graphConnected = $false
        $retryAttempt = 0
        
        while (-not $graphConnected -and $retryAttempt -lt $RetryCount) {
            try {
                if (-not $ForceReconnect -and -not $tokenExpired) {
                    # Check if already connected
                    $mgContext = Get-MgContext -ErrorAction Stop
                    if ($mgContext) {
                        Write-Log -Message "Already connected to Microsoft Graph as: $($mgContext.Account)" -Level "INFO"
                        $graphConnected = $true
                    }
                    else {
                        throw "Not connected to Microsoft Graph"
                    }
                }
                else {
                    throw "Need to establish new connection"
                }
            }
            catch {
                $retryAttempt++
                Write-Log -Message "Connecting to Microsoft Graph (Attempt $retryAttempt of $RetryCount)..." -Level "INFO"
                
                try {
                    # Disconnect if there's an existing connection
                    if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
                        Disconnect-MgGraph -ErrorAction SilentlyContinue
                    }
                    
                    # Prepare connection parameters
                    $graphParams = @{
                        Scopes = @("User.Read.All", "Directory.Read.All", "Organization.Read.All")
                    }
                    
                    # Add credential if provided
                    if ($Credential) {
                        # Convert PSCredential to username/password for Microsoft Graph
                        # This requires a different approach than Exchange Online
                        # The specific parameters depend on the Microsoft Graph PowerShell module version
                        $graphParams.Add('Credential', $Credential)
                    }
                    
                    # Connect to Microsoft Graph
                    Connect-MgGraph @graphParams
                    
                    $mgContext = Get-MgContext
                    $global:LastConnectionTime = Get-Date
                    $graphConnected = $true
                    
                    Write-Log -Message "Connected to Microsoft Graph as: $($mgContext.Account)" -Level "SUCCESS"
                }
                catch {
                    if ($retryAttempt -ge $RetryCount) {
                        Write-Log -Message "Failed to connect to Microsoft Graph after $RetryCount attempts: $_" -Level "ERROR" -ErrorCode "ERR005"
                        Write-Log -Message "Check your internet connection and credentials. Try manually running Connect-MgGraph in a separate window." -Level "ERROR"
                    }
                    else {
                        Write-Log -Message "Connection attempt failed, retrying in 5 seconds: $_" -Level "WARNING"
                        Start-Sleep -Seconds 5
                    }
                }
            }
        }
        
        # Verify migration endpoint if both connections succeeded
        if ($exchangeConnected -and $graphConnected) {
            try {
                $migrationEndpoint = Get-MigrationEndpoint -Identity $script:Config.MigrationEndpointName -ErrorAction Stop
                Write-Log -Message "Migration endpoint validated: $($migrationEndpoint.Identity)" -Level "SUCCESS"
                
                # Check endpoint type to ensure it's appropriate for hybrid migration
                if ($migrationEndpoint.EndpointType -ne "ExchangeRemoteMove") {
                    Write-Log -Message "Warning: Migration endpoint type is $($migrationEndpoint.EndpointType), but ExchangeRemoteMove is recommended for hybrid migration" -Level "WARNING"
                }
                
                # Test endpoint connectivity
                try {
                    $testResult = Test-MigrationServerAvailability -Endpoint $migrationEndpoint.Identity -ErrorAction Stop
                    Write-Log -Message "Migration endpoint connectivity test passed" -Level "SUCCESS"
                }
                catch {
                    Write-Log -Message "Migration endpoint connectivity test failed: $_" -Level "WARNING"
                    Write-Log -Message "The endpoint exists but may not be properly configured - check network connectivity" -Level "WARNING"
                }
                
                # Record connection metrics
                Record-MigrationMetric -MetricName "ConnectionEstablished" -Value "Success" -Properties @{
                    ExchangeConnected = $exchangeConnected
                    GraphConnected = $graphConnected
                    MigrationEndpoint = $migrationEndpoint.Identity
                    EndpointType = $migrationEndpoint.EndpointType
                }
            }
            catch {
                Write-Log -Message "Failed to validate migration endpoint: $_" -Level "ERROR" -ErrorCode "ERR006"
                Write-Log -Message "Available migration endpoints:" -Level "INFO"
                
                try {
                    $endpoints = Get-MigrationEndpoint
                    if ($endpoints.Count -eq 0) {
                        Write-Log -Message "  No migration endpoints found. You need to create one first." -Level "INFO"
                        Write-Log -Message "  Run 'New-MigrationEndpoint -ExchangeRemoteMove -Name \"Hybrid Migration Endpoint\" -RemoteServer [YourExchangeServerFQDN]'" -Level "INFO"
                    }
                    else {
                        foreach ($endpoint in $endpoints) {
                            Write-Log -Message "  - $($endpoint.Identity) ($($endpoint.EndpointType))" -Level "INFO"
                        }
                    }
                }
                catch {
                    Write-Log -Message "Could not retrieve available migration endpoints" -Level "WARNING"
                }
                
                # Record connection failure
                Record-MigrationMetric -MetricName "ConnectionEstablished" -Value "Failure" -Properties @{
                    ExchangeConnected = $exchangeConnected
                    GraphConnected = $graphConnected
                    FailureReason = "MigrationEndpointValidationFailed"
                    ErrorMessage = $_.Exception.Message
                }
                
                return $false
            }
        }
        else {
            # Record connection failure
            Record-MigrationMetric -MetricName "ConnectionEstablished" -Value "Failure" -Properties @{
                ExchangeConnected = $exchangeConnected
                GraphConnected = $graphConnected
                FailureReason = "ServiceConnectionFailed"
            }
            
            return $false
        }
        
        # Add a healthy connection check to ensure we have a good connection
        try {
            # Try a simple command to verify connection health
            $null = Get-AcceptedDomain -ErrorAction Stop
            $null = Get-MgUser -Top 1 -ErrorAction Stop
            
            Write-Log -Message "Connection health check passed for both Exchange Online and Microsoft Graph" -Level "SUCCESS"
            return $true
        }
        catch {
            Write-Log -Message "Connection health check failed: $_" -Level "ERROR"
            Write-Log -Message "Try reconnecting with -ForceReconnect parameter" -Level "ERROR"
            
            # Record health check failure
            Record-MigrationMetric -MetricName "ConnectionHealthCheck" -Value "Failure" -Properties @{
                ErrorMessage = $_.Exception.Message
            }
            
            return $false
        }
    }
    catch {
        Write-Log -Message "Failed to connect to migration services: $_" -Level "ERROR"
        
        # Record connection error
        Record-MigrationMetric -MetricName "ConnectionEstablished" -Value "Error" -Properties @{
            ErrorMessage = $_.Exception.Message
            ErrorType = $_.Exception.GetType().Name
        }
        
        return $false
    }
}