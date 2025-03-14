function Connect-MigrationEnvironment {
    <#
    .SYNOPSIS
        Connects to Exchange environments for migration operations.
    
    .DESCRIPTION
        Establishes authenticated connections to Exchange Online, on-premises Exchange,
        and Microsoft Graph as needed for migration operations. Supports multiple
        authentication methods and verifies connectivity to all required services.
    
    .PARAMETER Environment
        Specifies which environment(s) to connect to:
        - OnPremises: Connect only to on-premises Exchange
        - ExchangeOnline: Connect only to Exchange Online and Microsoft Graph
        - Both: Connect to both environments (default)
    
    .PARAMETER ForceReconnect
        When specified, forces a new connection even if already connected.
    
    .PARAMETER TokenLifetimeMinutes
        The expected lifetime of authentication tokens in minutes. Default is 50 minutes.
    
    .PARAMETER RetryCount
        Number of connection attempts before giving up. Default is 2.
    
    .PARAMETER Credential
        Optional PSCredential object containing authentication credentials.
    
    .PARAMETER CertificateThumbprint
        Certificate thumbprint for certificate-based authentication to Exchange Online.
    
    .PARAMETER ApplicationId
        Application ID for certificate-based or modern authentication to Exchange Online.
    
    .PARAMETER OrganizationName
        The on-premises Exchange organization name for connection.
    
    .PARAMETER OnPremisesExchangeServer
        The FQDN of the on-premises Exchange server to connect to.
    
    .EXAMPLE
        Connect-MigrationEnvironment -Environment ExchangeOnline
    
    .EXAMPLE
        Connect-MigrationEnvironment -Environment Both -Credential $cred
    
    .EXAMPLE
        Connect-MigrationEnvironment -Environment OnPremises -OnPremisesExchangeServer "exchange.contoso.local"
    
    .OUTPUTS
        [PSCustomObject] Returns an object with connection status information.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('OnPremises', 'ExchangeOnline', 'Both')]
        [string]$Environment = 'Both',
        
        [Parameter(Mandatory = $false)]
        [switch]$ForceReconnect,
        
        [Parameter(Mandatory = $false)]
        [int]$TokenLifetimeMinutes = 50,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 2,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint = "",
        
        [Parameter(Mandatory = $false)]
        [string]$ApplicationId = "",
        
        [Parameter(Mandatory = $false)]
        [string]$OrganizationName = "",
        
        [Parameter(Mandatory = $false)]
        [string]$OnPremisesExchangeServer = ""
    )
    
    # Initialize result object
    $result = [PSCustomObject]@{
        ExchangeOnlineConnected = $false
        OnPremisesConnected = $false
        GraphConnected = $false
        MigrationEndpointValidated = $false
        ValidationErrors = @()
    }
    
    try {
        # Track if we need to reconnect based on token lifetime
        $tokenExpired = $false
        
        if ($global:LastConnectionTime -ne $null) {
            $timeSinceLastConnection = (Get-Date) - $global:LastConnectionTime
            $tokenExpired = $timeSinceLastConnection.TotalMinutes -gt $TokenLifetimeMinutes
            
            if ($tokenExpired) {
                Write-Log -Message "Authentication token may be expired (last connection: $global:LastConnectionTime)" -Level "WARNING"
            }
        }
        
        # Get credential if not provided and needed for modern auth
        if (-not $Credential -and 
            ($Environment -ne 'OnPremises' -or [string]::IsNullOrEmpty($CertificateThumbprint)) -and
            [string]::IsNullOrEmpty($ApplicationId)) {
            try {
                $Credential = Get-MigrationCredential
            }
            catch {
                Write-Log -Message "Failed to get migration credentials: $_" -Level "WARNING"
                Write-Log -Message "Will attempt to connect using current context or device code authentication" -Level "INFO"
            }
        }
        
        # 1. Connect to Exchange Online and Microsoft Graph if needed
        if ($Environment -in @('ExchangeOnline', 'Both')) {
            # Connect to Exchange Online
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
                        
                        # Use certificate auth if specified
                        if (-not [string]::IsNullOrEmpty($CertificateThumbprint) -and -not [string]::IsNullOrEmpty($ApplicationId)) {
                            $exoParams.Add('CertificateThumbprint', $CertificateThumbprint)
                            $exoParams.Add('AppId', $ApplicationId)
                        }
                        # Otherwise add credential if provided
                        elseif ($Credential) {
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
                            $result.ValidationErrors += "Failed to connect to Exchange Online: $_"
                        }
                        else {
                            Write-Log -Message "Connection attempt failed, retrying in 5 seconds: $_" -Level "WARNING"
                            Start-Sleep -Seconds 5
                        }
                    }
                }
            }
            
            $result.ExchangeOnlineConnected = $exchangeConnected
            
            # Connect to Microsoft Graph if Exchange Online connected
            if ($exchangeConnected) {
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
                            
                            # Use certificate auth if specified
                            if (-not [string]::IsNullOrEmpty($CertificateThumbprint) -and -not [string]::IsNullOrEmpty($ApplicationId)) {
                                $graphParams.Add('CertificateThumbprint', $CertificateThumbprint)
                                $graphParams.Add('ClientId', $ApplicationId)
                            }
                            # Otherwise add credential if provided
                            elseif ($Credential) {
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
                                $result.ValidationErrors += "Failed to connect to Microsoft Graph: $_"
                            }
                            else {
                                Write-Log -Message "Connection attempt failed, retrying in 5 seconds: $_" -Level "WARNING"
                                Start-Sleep -Seconds 5
                            }
                        }
                    }
                }
                
                $result.GraphConnected = $graphConnected
                
                # Verify migration endpoint if both connections succeeded
                if ($exchangeConnected -and $graphConnected -and $script:Config.MigrationEndpointName) {
                    try {
                        $migrationEndpoint = Get-MigrationEndpoint -Identity $script:Config.MigrationEndpointName -ErrorAction Stop
                        Write-Log -Message "Migration endpoint validated: $($migrationEndpoint.Identity)" -Level "SUCCESS"
                        
                        # Check endpoint type to ensure it's appropriate for migration
                        if ($migrationEndpoint.EndpointType -ne "ExchangeRemoteMove") {
                            Write-Log -Message "Warning: Migration endpoint type is $($migrationEndpoint.EndpointType), but ExchangeRemoteMove is recommended for hybrid migration" -Level "WARNING"
                        }
                        
                        # Test endpoint connectivity
                        try {
                            $testResult = Test-MigrationServerAvailability -Endpoint $migrationEndpoint.Identity -ErrorAction Stop
                            Write-Log -Message "Migration endpoint connectivity test passed" -Level "SUCCESS"
                            $result.MigrationEndpointValidated = $true
                        }
                        catch {
                            Write-Log -Message "Migration endpoint connectivity test failed: $_" -Level "WARNING"
                            Write-Log -Message "The endpoint exists but may not be properly configured - check network connectivity" -Level "WARNING"
                            $result.ValidationErrors += "Migration endpoint connectivity test failed: $_"
                        }
                    }
                    catch {
                        Write-Log -Message "Failed to validate migration endpoint: $_" -Level "ERROR" -ErrorCode "ERR006"
                        $result.ValidationErrors += "Failed to validate migration endpoint: $_"
                        
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
                    }
                }
            }
        }
        
        # 2. Connect to on-premises Exchange if needed
        if ($Environment -in @('OnPremises', 'Both')) {
            $onPremConnected = $false
            $retryAttempt = 0
            
            while (-not $onPremConnected -and $retryAttempt -lt $RetryCount) {
                try {
                    if (-not $ForceReconnect) {
                        # Check if already connected to on-premises Exchange
                        $onPremServer = Get-ExchangeServer -ErrorAction Stop
                        if ($onPremServer) {
                            Write-Log -Message "Already connected to on-premises Exchange: $($onPremServer[0].Name)" -Level "INFO"
                            $onPremConnected = $true
                        }
                    }
                    else {
                        throw "Need to establish new connection"
                    }
                }
                catch {
                    $retryAttempt++
                    Write-Log -Message "Connecting to on-premises Exchange (Attempt $retryAttempt of $RetryCount)..." -Level "INFO"
                    
                    try {
                        # Check for existing session and remove if necessary
                        $existingSession = Get-PSSession | Where-Object { 
                            $_.ConfigurationName -eq "Microsoft.Exchange" -and 
                            $_.ComputerName -ne "outlook.office365.com" -and
                            $_.State -eq "Opened" 
                        }
                        
                        if ($existingSession) {
                            Write-Log -Message "Removing existing on-premises Exchange session" -Level "INFO"
                            $existingSession | Remove-PSSession -ErrorAction SilentlyContinue
                        }
                        
                        # Prepare connection parameters
                        $onPremParams = @{
                            ConnectionUri = if ($OnPremisesExchangeServer) { 
                                "http://$OnPremisesExchangeServer/PowerShell/" 
                            } else { 
                                $script:Config.OnPremisesExchangeUri 
                            }
                            Authentication = "Kerberos"
                        }
                        
                        # Add organization if specified
                        if (-not [string]::IsNullOrEmpty($OrganizationName)) {
                            $onPremParams.Add('Organization', $OrganizationName)
                        }
                        elseif (-not [string]::IsNullOrEmpty($script:Config.Organization)) {
                            $onPremParams.Add('Organization', $script:Config.Organization)
                        }
                        
                        # Add credential if provided
                        if ($Credential) {
                            $onPremParams.Add('Credential', $Credential)
                        }
                        
                        # Connect to on-premises Exchange
                        $session = New-PSSession @onPremParams
                        Import-PSSession $session -DisableNameChecking -AllowClobber
                        
                        $onPremServer = Get-ExchangeServer
                        $onPremConnected = $true
                        
                        Write-Log -Message "Connected to on-premises Exchange: $($onPremServer[0].Name)" -Level "SUCCESS"
                    }
                    catch {
                        if ($retryAttempt -ge $RetryCount) {
                            Write-Log -Message "Failed to connect to on-premises Exchange after $RetryCount attempts: $_" -Level "ERROR" -ErrorCode "ERR007"
                            Write-Log -Message "Check your network connection and credentials. Try manually connecting to the Exchange PowerShell URL." -Level "ERROR"
                            $result.ValidationErrors += "Failed to connect to on-premises Exchange: $_"
                        }
                        else {
                            Write-Log -Message "Connection attempt failed, retrying in 5 seconds: $_" -Level "WARNING"
                            Start-Sleep -Seconds 5
                        }
                    }
                }
            }
            
            $result.OnPremisesConnected = $onPremConnected
        }
        
        # Final status check
        $allRequiredConnected = $true
        
        if ($Environment -eq 'ExchangeOnline' -and (-not $result.ExchangeOnlineConnected -or -not $result.GraphConnected)) {
            $allRequiredConnected = $false
        }
        elseif ($Environment -eq 'OnPremises' -and -not $result.OnPremisesConnected) {
            $allRequiredConnected = $false
        }
        elseif ($Environment -eq 'Both' -and (-not $result.ExchangeOnlineConnected -or -not $result.OnPremisesConnected)) {
            $allRequiredConnected = $false
        }
        
        if ($allRequiredConnected) {
            Write-Log -Message "Successfully connected to all required migration environments" -Level "SUCCESS"
            
            # Record connection metrics
            Record-MigrationMetric -MetricName "ConnectionEstablished" -Value "Success" -Properties @{
                ExchangeOnlineConnected = $result.ExchangeOnlineConnected
                OnPremisesConnected = $result.OnPremisesConnected
                GraphConnected = $result.GraphConnected
                Environment = $Environment
            }
        }
        else {
            Write-Log -Message "Failed to connect to one or more required migration environments" -Level "ERROR"
            
            # Record connection failure
            Record-MigrationMetric -MetricName "ConnectionEstablished" -Value "Failure" -Properties @{
                ExchangeOnlineConnected = $result.ExchangeOnlineConnected
                OnPremisesConnected = $result.OnPremisesConnected
                GraphConnected = $result.GraphConnected
                Environment = $Environment
                ValidationErrors = ($result.ValidationErrors -join "; ")
            }
        }
        
        return $result
    }
    catch {
        Write-Log -Message "Unexpected error in Connect-MigrationEnvironment: $_" -Level "ERROR"
        
        # Record connection error
        Record-MigrationMetric -MetricName "ConnectionEstablished" -Value "Error" -Properties @{
            ErrorMessage = $_.Exception.Message
            ErrorType = $_.Exception.GetType().Name
            Environment = $Environment
        }
        
        # Add error to result and return
        $result.ValidationErrors += "Unexpected error: $_"
        return $result
    }
}
