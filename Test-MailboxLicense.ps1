function Test-MailboxLicense {
    <#
    .SYNOPSIS
        Tests if a mailbox has the appropriate license for migration.
    
    .DESCRIPTION
        Checks if the specified mailbox has an Exchange Online license assigned and properly provisioned.
        Validates if E1 or E5 licenses are present and if Exchange services are enabled.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MailboxLicense -EmailAddress "user@contoso.com" -Results $results
    
    .OUTPUTS
        [bool] Returns $true if the test was successful (even if issues were found), $false if the test failed.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        # Get user license info from Microsoft Graph
        try {
            $mgUser = Get-MgUser -UserId $EmailAddress -Property DisplayName, UserPrincipalName, UsageLocation, Id, ProxyAddresses -ErrorAction Stop
        }
        catch [Microsoft.Graph.PowerShell.Authentication.Models.AuthenticationException] {
            Write-Log -Message "Graph API authentication error: $($_.Exception.Message)" -Level "ERROR"
            $Results.Errors += "Failed to authenticate with Microsoft Graph API to check licenses"
            return $false
        }
        catch [Microsoft.Graph.PowerShell.Models.ODataErrors.ODataError] {
            if ($_.Exception.Message -like "*Resource '*' does not exist*") {
                Write-Log -Message "User not found in Microsoft Graph API: $EmailAddress" -Level "ERROR"
                $Results.Errors += "User $EmailAddress not found in Azure AD. Verify the user exists and is synced."
                return $false
            }
            else {
                Write-Log -Message "Graph API error retrieving user: $($_.Exception.Message)" -Level "ERROR"
                $Results.Errors += "Graph API error: $($_.Exception.Message)"
                return $false
            }
        }
        catch {
            Write-Log -Message "Failed to get Microsoft Graph user: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
            $Results.Errors += "Failed to get user information from Microsoft Graph: $($_.Exception.Message)"
            return $false
        }
        
        try {
            $mgUserLicense = Get-MgUserLicenseDetail -UserId $EmailAddress -ErrorAction Stop
        }
        catch [Microsoft.Graph.PowerShell.Models.ODataErrors.ODataError] {
            Write-Log -Message "Graph API error retrieving license details: $($_.Exception.Message)" -Level "WARNING"
            $Results.Warnings += "Unable to retrieve license details from Graph API: $($_.Exception.Message)"
            $mgUserLicense = $null
        }
        catch {
            Write-Log -Message "Failed to get license details: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
            $Results.Warnings += "Failed to get license details: $($_.Exception.Message)"
            $mgUserLicense = $null
        }
        
        if ($null -eq $mgUserLicense) {
            $Results.Errors += "No licenses assigned to user"
            $Results.ErrorCodes += "ERR012"
            Write-Log -Message "Error: No licenses assigned to $EmailAddress" -Level "ERROR" -ErrorCode "ERR012"
            Write-Log -Message "Troubleshooting: Assign an Exchange Online license to the user" -Level "INFO"
            return $true
        }
        
        # Modified license check: First check for E1 or E5 licenses
        $e1License = $mgUserLicense | Where-Object { 
            $_.SkuPartNumber -eq "STANDARDPACK" # E1
        }
        
        $e5License = $mgUserLicense | Where-Object { 
            $_.SkuPartNumber -eq "ENTERPRISEPACK" # E5
        }
        
        $e3License = $mgUserLicense | Where-Object { 
            $_.SkuPartNumber -eq "ENTERPRISEPACKPLUS" # E3
        }
        
        # Set license type based on what's found
        if ($e5License) {
            $Results.HasE1OrE5License = $true
            $Results.LicenseType = "E5"
            
            # Now check if Exchange is enabled within the E5 license
            $exchangeService = $e5License.ServicePlans | Where-Object {
                $_.ServicePlanName -eq "EXCHANGE_S_ENTERPRISE"
            }
            
            if ($exchangeService) {
                $Results.HasExchangeLicense = $true
                $Results.LicenseDetails = "EXCHANGE_S_ENTERPRISE"
                $Results.LicenseProvisioningStatus = $exchangeService.ProvisioningStatus
                
                if ($exchangeService.ProvisioningStatus -eq "Error") {
                    $Results.Errors += "E5 Exchange license provisioning error"
                    $Results.ErrorCodes += "ERR013"
                    Write-Log -Message "Error: E5 Exchange license provisioning error for $EmailAddress" -Level "ERROR" -ErrorCode "ERR013"
                    Write-Log -Message "Troubleshooting: Review the service health in Microsoft 365 admin center or try removing and reassigning the license." -Level "INFO"
                }
                elseif ($exchangeService.ProvisioningStatus -ne "Success") {
                    $Results.Warnings += "E5 Exchange license not fully provisioned: $($exchangeService.ProvisioningStatus)"
                    Write-Log -Message "Warning: E5 Exchange license not fully provisioned for $EmailAddress`: $($exchangeService.ProvisioningStatus)" -Level "WARNING"
                }
            }
            else {
                $Results.Errors += "E5 license found but Exchange Online service not enabled"
                $Results.ErrorCodes += "ERR012"
                Write-Log -Message "Error: E5 license found but Exchange Online service not enabled for $EmailAddress" -Level "ERROR" -ErrorCode "ERR012"
                Write-Log -Message "Troubleshooting: Verify the Exchange Online service is enabled in the license options" -Level "INFO"
            }
        }
        elseif ($e3License) {
            $Results.HasE1OrE5License = $true  # E3 is fine too
            $Results.LicenseType = "E3"
            
            # Check if Exchange is enabled within the E3 license
            $exchangeService = $e3License.ServicePlans | Where-Object {
                $_.ServicePlanName -eq "EXCHANGE_S_ENTERPRISE"
            }
            
            if ($exchangeService) {
                $Results.HasExchangeLicense = $true
                $Results.LicenseDetails = "EXCHANGE_S_ENTERPRISE"
                $Results.LicenseProvisioningStatus = $exchangeService.ProvisioningStatus
                
                if ($exchangeService.ProvisioningStatus -eq "Error") {
                    $Results.Errors += "E3 Exchange license provisioning error"
                    $Results.ErrorCodes += "ERR013"
                    Write-Log -Message "Error: E3 Exchange license provisioning error for $EmailAddress" -Level "ERROR" -ErrorCode "ERR013"
                    Write-Log -Message "Troubleshooting: Review the service health in Microsoft 365 admin center or try removing and reassigning the license." -Level "INFO"
                }
                elseif ($exchangeService.ProvisioningStatus -ne "Success") {
                    $Results.Warnings += "E3 Exchange license not fully provisioned: $($exchangeService.ProvisioningStatus)"
                    Write-Log -Message "Warning: E3 Exchange license not fully provisioned for $EmailAddress`: $($exchangeService.ProvisioningStatus)" -Level "WARNING"
                }
            }
            else {
                $Results.Errors += "E3 license found but Exchange Online service not enabled"
                $Results.ErrorCodes += "ERR012"
                Write-Log -Message "Error: E3 license found but Exchange Online service not enabled for $EmailAddress" -Level "ERROR" -ErrorCode "ERR012"
                Write-Log -Message "Troubleshooting: Verify the Exchange Online service is enabled in the license options" -Level "INFO"
            }
        }
        elseif ($e1License) {
            $Results.HasE1OrE5License = $true
            $Results.LicenseType = "E1"
            
            # Now check if Exchange is enabled within the E1 license
            $exchangeService = $e1License.ServicePlans | Where-Object {
                $_.ServicePlanName -eq "EXCHANGE_S_STANDARD"
            }
            
            if ($exchangeService) {
                $Results.HasExchangeLicense = $true
                $Results.LicenseDetails = "EXCHANGE_S_STANDARD"
                $Results.LicenseProvisioningStatus = $exchangeService.ProvisioningStatus
                
                if ($exchangeService.ProvisioningStatus -eq "Error") {
                    $Results.Errors += "E1 Exchange license provisioning error"
                    $Results.ErrorCodes += "ERR013"
                    Write-Log -Message "Error: E1 Exchange license provisioning error for $EmailAddress" -Level "ERROR" -ErrorCode "ERR013"
                    Write-Log -Message "Troubleshooting: Review the service health in Microsoft 365 admin center or try removing and reassigning the license." -Level "INFO"
                }
                elseif ($exchangeService.ProvisioningStatus -ne "Success") {
                    $Results.Warnings += "E1 Exchange license not fully provisioned: $($exchangeService.ProvisioningStatus)"
                    Write-Log -Message "Warning: E1 Exchange license not fully provisioned for $EmailAddress`: $($exchangeService.ProvisioningStatus)" -Level "WARNING"
                }
            }
            else {
                $Results.Errors += "E1 license found but Exchange Online service not enabled"
                $Results.ErrorCodes += "ERR012"
                Write-Log -Message "Error: E1 license found but Exchange Online service not enabled for $EmailAddress" -Level "ERROR" -ErrorCode "ERR012"
                Write-Log -Message "Troubleshooting: Verify the Exchange Online service is enabled in the license options" -Level "INFO"
            }
        }
        else {
            # Fallback to check for any Exchange license if neither E1 nor E5 is found
            $mgUserExchangeLicense = $mgUserLicense.ServicePlans | Where-Object { 
                $_.ServicePlanName -like 'EXCHANGE_S_*' -and $_.AppliesTo -eq 'User'
            }
            
            if ($mgUserExchangeLicense) {
                $Results.HasExchangeLicense = $true
                $Results.LicenseDetails = $mgUserExchangeLicense.ServicePlanName -join ','
                $Results.LicenseProvisioningStatus = $mgUserExchangeLicense.ProvisioningStatus -join ','
                
                if ($script:Config.RequiredE1orE5License) {
                    $Results.Warnings += "Exchange license found but not E1, E3, or E5: $($mgUserExchangeLicense.ServicePlanName -join ',')"
                    Write-Log -Message "Warning: Exchange license found but not E1, E3, or E5 for $EmailAddress`: $($mgUserExchangeLicense.ServicePlanName -join ',')" -Level "WARNING"
                }
                
                if ($mgUserExchangeLicense.ProvisioningStatus -contains "Error") {
                    $Results.Errors += "License provisioning error"
                    $Results.ErrorCodes += "ERR013"
                    Write-Log -Message "Error: License provisioning error for $EmailAddress" -Level "ERROR" -ErrorCode "ERR013"
                    Write-Log -Message "Troubleshooting: Review the service health in Microsoft 365 admin center or try removing and reassigning the license." -Level "INFO"
                }
                elseif ($mgUserExchangeLicense.ProvisioningStatus -notcontains "Success") {
                    $Results.Warnings += "License not fully provisioned: $($mgUserExchangeLicense.ProvisioningStatus -join ',')"
                    Write-Log -Message "Warning: License not fully provisioned for $EmailAddress`: $($mgUserExchangeLicense.ProvisioningStatus -join ',')" -Level "WARNING"
                }
            }
            else {
                $Results.Errors += "No Exchange Online license assigned"
                $Results.ErrorCodes += "ERR012"
                Write-Log -Message "Error: No Exchange Online license assigned to $EmailAddress" -Level "ERROR" -ErrorCode "ERR012"
                Write-Log -Message "Troubleshooting: Assign an Exchange Online license (preferably E1, E3, or E5) to the user" -Level "INFO"
            }
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Unexpected error checking license: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
        if ($_.ScriptStackTrace) {
            Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
        }
        $Results.Errors += "Failed to get user license info: $_"
        return $false
    }
}
