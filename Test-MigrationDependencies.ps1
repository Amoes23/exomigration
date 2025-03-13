function Test-MigrationDependencies {
    <#
    .SYNOPSIS
        Checks if all required PowerShell modules are installed.
    
    .DESCRIPTION
        Verifies that all necessary PowerShell modules are installed with the required
        minimum versions and provides installation instructions if any are missing.
    
    .EXAMPLE
        Test-MigrationDependencies
    
    .OUTPUTS
        [bool] True if all dependencies are met, False otherwise.
    #>
    [CmdletBinding()]
    param()
    
    Write-Log -Message "Checking for required PowerShell modules..." -Level "INFO"
    
    $requiredModules = @(
        @{
            Name = "ExchangeOnlineManagement"
            MinimumVersion = "3.0.0"
            Description = "Exchange Online PowerShell V3 module"
            InstallCommand = "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber"
        },
        @{
            Name = "Microsoft.Graph"
            MinimumVersion = "1.20.0" 
            Description = "Microsoft Graph PowerShell SDK"
            InstallCommand = "Install-Module -Name Microsoft.Graph -Force -AllowClobber"
        },
        @{
            Name = "Microsoft.Graph.Users"
            MinimumVersion = "1.20.0"
            Description = "Microsoft Graph Users module"
            InstallCommand = "Install-Module -Name Microsoft.Graph.Users -Force -AllowClobber"
        }
    )
    
    $missingModules = @()
    $outdatedModules = @()
    
    foreach ($module in $requiredModules) {
        $installedModule = Get-Module -Name $module.Name -ListAvailable
        
        if (-not $installedModule) {
            $missingModules += $module
            Write-Log -Message "Required module not found: $($module.Name) - $($module.Description)" -Level "ERROR" -ErrorCode "ERR001"
        }
        else {
            $latestVersion = ($installedModule | Sort-Object Version -Descending)[0].Version
            
            if ($latestVersion -lt [Version]$module.MinimumVersion) {
                $outdatedModules += @{
                    Module = $module
                    InstalledVersion = $latestVersion
                }
                Write-Log -Message "Module version too old: $($module.Name) - Current: $latestVersion, Required: $($module.MinimumVersion)" -Level "WARNING"
            }
            else {
                Write-Log -Message "Required module found: $($module.Name) (v$latestVersion)" -Level "DEBUG"
            }
        }
    }
    
    if ($missingModules.Count -gt 0 -or $outdatedModules.Count -gt 0) {
        Write-Log -Message "Dependency issues found. Please resolve before continuing." -Level "ERROR" -ErrorCode "ERR001"
        
        if ($missingModules.Count -gt 0) {
            Write-Log -Message "Missing modules:" -Level "ERROR"
            foreach ($module in $missingModules) {
                Write-Log -Message "- $($module.Name): $($module.Description)" -Level "ERROR"
                Write-Log -Message "  Install command: $($module.InstallCommand)" -Level "INFO"
            }
        }
        
        if ($outdatedModules.Count -gt 0) {
            Write-Log -Message "Outdated modules:" -Level "WARNING"
            foreach ($moduleInfo in $outdatedModules) {
                Write-Log -Message "- $($moduleInfo.Module.Name): Current v$($moduleInfo.InstalledVersion), Required v$($moduleInfo.Module.MinimumVersion)" -Level "WARNING"
                Write-Log -Message "  Update command: $($moduleInfo.Module.InstallCommand)" -Level "INFO"
            }
        }
        
        return $false
    }
    
    # Check PowerShell version recommendation
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        Write-Log -Message "PowerShell 7+ is recommended for optimal performance. Current version: $($PSVersionTable.PSVersion)" -Level "WARNING"
    }
    
    # Check for permission to run commands
    try {
        # Try to execute a simple command to verify we can execute commands
        $null = Get-PSCallStack
        
        Write-Log -Message "Permission check passed" -Level "DEBUG"
    }
    catch {
        Write-Log -Message "Insufficient permissions to run PowerShell commands. Try running as Administrator." -Level "ERROR"
        return $false
    }
    
    # Check for internet connectivity
    try {
        $connected = Test-NetConnection -ComputerName outlook.office365.com -Port 443 -InformationLevel Quiet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        
        if (-not $connected) {
            Write-Log -Message "No internet connectivity to Exchange Online detected. Check your network connection." -Level "ERROR"
            return $false
        }
        
        Write-Log -Message "Internet connectivity check passed" -Level "DEBUG"
    }
    catch {
        Write-Log -Message "Failed to check internet connectivity: $_" -Level "WARNING"
        Write-Log -Message "Make sure you have internet access to connect to Exchange Online" -Level "WARNING"
    }
    
    Write-Log -Message "All required dependencies are installed" -Level "SUCCESS"
    return $true
}