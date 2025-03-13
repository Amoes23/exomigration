function Test-ADSyncStatus {
    <#
    .SYNOPSIS
        Tests the Active Directory synchronization status for a mailbox.
    
    .DESCRIPTION
        Verifies that the mailbox object is properly synchronized between on-premises
        Active Directory and Azure AD, which is critical for hybrid migrations.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-ADSyncStatus -EmailAddress "user@contoso.com" -Results $results
    
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
    
    if (-not ($script:Config.CheckADSyncStatus)) {
        Write-Log -Message "Skipping AD sync status check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        Write-Log -Message "Checking AD sync status for: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "ADSyncVerified" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "IsDirSynced" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "LastDirSyncTime" -NotePropertyValue $null -Force
        $Results | Add-Member -NotePropertyName "SyncIssues" -NotePropertyValue @() -Force
        
        # Get Exchange Online mailbox information
        try {
            $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
            
            # Check if the mailbox is DirSynced
            $Results.IsDirSynced = $mailbox.IsDirSynced -eq $true
        }
        catch [Microsoft.Exchange.Management.PowerShell.CommandNotFoundException] {
            Write-Log -Message "Exchange Online connection issue: Get-Mailbox cmdlet not found" -Level "ERROR"
            $Results.SyncIssues += "Unable to verify Exchange Online connection. Check if you're connected to Exchange Online."
            return $false
        }
        catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
            Write-Log -Message "Insufficient permissions to access mailbox: $EmailAddress" -Level "ERROR"
            $Results.SyncIssues += "Insufficient permissions to access mailbox information. Check your Exchange Online permissions."
            return $false
        }
        catch [System.Management.Automation.RemoteException] {
            Write-Log -Message "Remote Exchange error: $($_.Exception.Message)" -Level "ERROR"
            $Results.SyncIssues += "Exchange Online remote error: $($_.Exception.Message)"
            return $false
        }
        catch {
            Write-Log -Message "Failed to get mailbox: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
            $Results.SyncIssues += "Failed to get mailbox information: $($_.Exception.Message)"
            return $false
        }
        
        # Get user from Microsoft Graph for more detailed sync information
        try {
            $mgUser = Get-MgUser -UserId $EmailAddress -ErrorAction Stop
            
            # Check for onPremisesSyncEnabled
            $onPremisesSyncEnabled = $mgUser.OnPremisesSyncEnabled -eq $true
            
            # Check for immutableId (source anchor)
            $hasImmutableId = -not [string]::IsNullOrEmpty($mgUser.OnPremisesImmutableId)
            
            # Get last sync time if available
            if ($mgUser.OnPremisesLastSyncDateTime) {
                $Results.LastDirSyncTime = $mgUser.OnPremisesLastSyncDateTime
                $syncTimeOk = ($mgUser.OnPremisesLastSyncDateTime -gt (Get-Date).AddDays(-1))
            }
            else {
                $syncTimeOk = $false
            }
            
            # Determine if sync is verified
            $Results.ADSyncVerified = $Results.IsDirSynced -and $onPremisesSyncEnabled -and $hasImmutableId -and $syncTimeOk
            
            # Collect any sync issues
            if (-not $Results.IsDirSynced) {
                $Results.SyncIssues += "Mailbox is not marked as DirSynced in Exchange Online"
            }
            
            if (-not $onPremisesSyncEnabled) {
                $Results.SyncIssues += "OnPremisesSyncEnabled is not set to true in Azure AD"
            }
            
            if (-not $hasImmutableId) {
                $Results.SyncIssues += "Mailbox is missing OnPremisesImmutableId (source anchor)"
            }
            
            if (-not $syncTimeOk) {
                if ($Results.LastDirSyncTime) {
                    $daysSinceSync = [math]::Round(((Get-Date) - $Results.LastDirSyncTime).TotalDays, 1)
                    $Results.SyncIssues += "Last directory sync was $daysSinceSync days ago (should be less than 1 day)"
                }
                else {
                    $Results.SyncIssues += "No last sync time found, object may not be syncing"
                }
            }
            
            # Add appropriate warnings or errors based on sync status
            if (-not $Results.ADSyncVerified) {
                $Results.Errors += "AD sync verification failed - mailbox may not be properly synchronized"
                $Results.ErrorCodes += "ERR030"
                
                Write-Log -Message "Error: AD sync verification failed for $EmailAddress" -Level "ERROR" -ErrorCode "ERR030"
                foreach ($issue in $Results.SyncIssues) {
                    Write-Log -Message "  - $issue" -Level "ERROR"
                }
                
                Write-Log -Message "Recommendation: Verify Azure AD Connect setup and force a sync" -Level "INFO"
            }
            else {
                Write-Log -Message "AD sync verified for $EmailAddress" -Level "SUCCESS"
                Write-Log -Message "Last directory sync: $($Results.LastDirSyncTime)" -Level "INFO"
            }
        }
        catch [Microsoft.Graph.PowerShell.Authentication.Models.AuthenticationException] {
            Write-Log -Message "Graph API authentication error: $($_.Exception.Message)" -Level "ERROR"
            $Results.SyncIssues += "Microsoft Graph authentication error. Verify your Graph API connection."
            $Results.Errors += "Failed to authenticate with Microsoft Graph API"
            $Results.ErrorCodes += "ERR031"
            return $false
        }
        catch [Microsoft.Graph.PowerShell.Models.ODataErrors.ODataError] {
            if ($_.Exception.Message -like "*Resource '*' does not exist*") {
                Write-Log -Message "User not found in Microsoft Graph API: $EmailAddress" -Level "ERROR" -ErrorCode "ERR031"
                $Results.SyncIssues += "User not found in Microsoft Graph API - check if the user exists in Azure AD"
                $Results.Errors += "User not found in Microsoft Graph API - check if the user exists in Azure AD"
                $Results.ErrorCodes += "ERR031"
            }
            else {
                Write-Log -Message "Graph API error: $($_.Exception.Message)" -Level "ERROR"
                $Results.SyncIssues += "Microsoft Graph API error: $($_.Exception.Message)"
                $Results.Errors += "Microsoft Graph API error: $($_.Exception.Message)"
                $Results.ErrorCodes += "ERR031"
            }
            return $false
        }
        catch {
            Write-Log -Message "Failed to check AD sync status: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
            $Results.SyncIssues += "Failed to check AD sync status: $($_.Exception.Message)"
            $Results.Warnings += "Failed to check AD sync status: $($_.Exception.Message)"
            return $false
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Unexpected error during AD sync check: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
        if ($_.ScriptStackTrace) {
            Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
        }
        $Results.Warnings += "Failed to check AD sync status: $_"
        return $false
    }
}
