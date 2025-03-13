function Import-MigrationConfig {
    <#
    .SYNOPSIS
        Imports configuration settings from a JSON file.
    
    .DESCRIPTION
        Loads migration configuration settings from the specified JSON file and merges them with default settings.
        Command-line parameters can override settings from the configuration file.
    
    .PARAMETER ConfigPath
        Path to the JSON configuration file. If the file doesn't exist, a default configuration file will be created.
    
    .PARAMETER OverrideParams
        Optional hashtable of parameters that will override settings from the configuration file.
    
    .EXAMPLE
        $config = Import-MigrationConfig -ConfigPath "C:\Migration\Config.json"
    
    .EXAMPLE
        $config = Import-MigrationConfig -ConfigPath "C:\Migration\Config.json" -OverrideParams @{TargetDeliveryDomain = "contoso.mail.onmicrosoft.com"}
    
    .OUTPUTS
        [hashtable] Configuration settings as a hashtable
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$OverrideParams = @{}
    )
    
    try {
        Write-Log -Message "Importing configuration from: $ConfigPath" -Level "INFO"
        
        # Define default configuration
        $defaultConfig = @{
            TargetDeliveryDomain = ""
            MigrationEndpointName = ""
            LogPath = ".\Logs\"
            ReportPath = ".\Reports\"
            NotificationEmails = @()
            CompleteAfterDays = 1
            MaxMailboxSizeGB = 50
            StartAfterMinutes = 15
            RequiredE1orE5License = $true
            CheckDomainVerification = $true
            BatchCreationTimeoutMinutes = 10
            HTMLTemplatePath = ".\Templates\ReportTemplate.html"
            RequireUPNMatchSMTP = $true
            TokenLifetimeMinutes = 50
            TokenRefreshBuffer = 5
            MaxItemSizeMB = 150
            ReportVersioning = $true
            DiskSpaceThresholdMB = 500
            EnableVerboseLogging = $false
            LogCleanupDays = 30
            ConnectionRetryCount = 3
            ConnectionRetryDelaySeconds = 5
            InactivityThresholdDays = 30
            MoveRequestDetailLevel = "Full"
            CheckIncompleteMailboxMoves = $true
            UseBadItemLimitRecommendations = $true
            ValidationLevel = "Standard"
            EnableCheckpointing = $true
            WorkspacePath = ".\Workspace\"
            EnableParallelProcessing = $true
            MaxConcurrentMailboxes = 5
            BatchSize = 100
            UseExternalTemplate = $true
            EnableMemoryOptimization = $true
            AuthenticationMethod = "Modern"
            CertificateThumbprint = ""
            ApplicationId = ""
            TenantId = ""
            Organization = ""
			OnPremisesExchangeUri = "" 
			CheckExchangeOnlineForConflicts = $true 
        }
        
        # Check if config file exists
        if (-not (Test-Path -Path $ConfigPath)) {
            # If config file doesn't exist, create directory and default config
            $configDir = Split-Path -Path $ConfigPath -Parent
            
            if (-not (Test-Path -Path $configDir)) {
                New-Item -ItemType Directory -Path $configDir -Force | Out-Null
                Write-Log -Message "Created config directory: $configDir" -Level "INFO"
            }
            
            # Create default config file
            $defaultConfig | ConvertTo-Json -Depth 10 | Out-File -FilePath $ConfigPath -Encoding utf8
            Write-Log -Message "Created default configuration file: $ConfigPath" -Level "INFO"
            
            # Return default config with overrides
            $configHashtable = $defaultConfig.Clone()
        }
        else {
            # Read and parse the JSON config file
            $configJson = Get-Content -Path $ConfigPath -Raw
            $config = $configJson | ConvertFrom-Json
            
            # Convert to hashtable
            $configHashtable = @{}
            
            # Get all default config properties and ensure they exist in loaded config
            foreach ($key in $defaultConfig.Keys) {
                if (Get-Member -InputObject $config -Name $key -MemberType Properties) {
                    $configHashtable[$key] = $config.$key
                }
                else {
                    # Use default value if missing in config file
                    $configHashtable[$key] = $defaultConfig[$key]
                    Write-Log -Message "Config file missing property '$key', using default value: $($defaultConfig[$key])" -Level "WARNING"
                }
            }
        }
        
        # Apply overrides from parameters
        foreach ($key in $OverrideParams.Keys) {
            if ($configHashtable.ContainsKey($key)) {
                $configHashtable[$key] = $OverrideParams[$key]
                Write-Log -Message "Using override for $($key): $($OverrideParams[$key])" -Level "INFO"
            }
        }
        
        # Validate critical configuration items
        if ([string]::IsNullOrWhiteSpace($configHashtable['TargetDeliveryDomain'])) {
            Write-Log -Message "TargetDeliveryDomain is missing in config and not provided as parameter" -Level "ERROR" -ErrorCode "ERR002"
            return $null
        }
        
        if ([string]::IsNullOrWhiteSpace($configHashtable['MigrationEndpointName'])) {
            Write-Log -Message "MigrationEndpointName is missing in config and not provided as parameter" -Level "ERROR" -ErrorCode "ERR002"
            return $null
        }
        
        if ($configHashtable['NotificationEmails'].Count -eq 0) {
            Write-Log -Message "No NotificationEmails specified in config or as parameter" -Level "WARNING"
        }
        
        Write-Log -Message "Configuration loaded successfully." -Level "SUCCESS"
        return $configHashtable
    }
    catch {
        Write-Log -Message "Failed to load configuration: $_" -Level "ERROR" -ErrorCode "ERR002"
        return $null
    }
}
