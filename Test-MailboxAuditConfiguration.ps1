function Test-MailboxAuditConfiguration {
    <#
    .SYNOPSIS
        Tests mailbox audit log configuration for migration readiness.
    
    .DESCRIPTION
        Checks if a mailbox has custom audit log settings that need to be preserved
        during migration to Exchange Online. Identifies audit configurations that
        might need to be reconfigured after migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MailboxAuditConfiguration -EmailAddress "user@contoso.com" -Results $results
    
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
    
    if (-not ($script:Config.CheckForAuditLogConfiguration)) {
        Write-Log -Message "Skipping audit log configuration check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        Write-Log -Message "Checking mailbox audit configuration for: $EmailAddress" -Level "INFO"
        
        # Get mailbox audit settings
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Initialize audit properties
        $Results | Add-Member -NotePropertyName "AuditEnabled" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "AuditLogAgeLimit" -NotePropertyValue 90 -Force  # Default is 90 days
        $Results | Add-Member -NotePropertyName "AuditOwner" -NotePropertyValue @() -Force
        $Results | Add-Member -NotePropertyName "AuditAdmin" -NotePropertyValue @() -Force
        $Results | Add-Member -NotePropertyName "AuditDelegate" -NotePropertyValue @() -Force
        $Results | Add-Member -NotePropertyName "CustomAuditSettings" -NotePropertyValue $false -Force
        
        # Check if audit logging is enabled for the mailbox
        $Results.AuditEnabled = $mailbox.AuditEnabled
        
        # Check audit log age limit
        if ($mailbox.AuditLogAgeLimit) {
            $Results.AuditLogAgeLimit = $mailbox.AuditLogAgeLimit.Days
        }
        
        # Check audit action settings
        if ($mailbox.AuditOwnerRules) {
            $Results.AuditOwner = $mailbox.AuditOwner
        }
        
        if ($mailbox.AuditAdminRules) {
            $Results.AuditAdmin = $mailbox.AuditAdmin
        }
        
        if ($mailbox.AuditDelegateRules) {
            $Results.AuditDelegate = $mailbox.AuditDelegate
        }
        
        # Check for custom audit settings
        $defaultOwnerActions = @("HardDelete", "SoftDelete", "Move", "MoveToDeletedItems")
        $defaultAdminActions = @("Update", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete")
        $defaultDelegateActions = @("Update", "Move", "MoveToDeletedItems", "SoftDelete", "HardDelete")
        
        $ownerCustomized = $false
        $adminCustomized = $false
        $delegateCustomized = $false
        
        # Check if owner actions are customized
        if ($Results.AuditOwner.Count -gt 0) {
            $ownerCustomized = -not ($Results.AuditOwner.Count -eq $defaultOwnerActions.Count -and 
                ($Results.AuditOwner | Where-Object { $defaultOwnerActions -notcontains $_ }).Count -eq 0)
        }
        
        # Check if admin actions are customized
        if ($Results.AuditAdmin.Count -gt 0) {
            $adminCustomized = -not ($Results.AuditAdmin.Count -eq $defaultAdminActions.Count -and 
                ($Results.AuditAdmin | Where-Object { $defaultAdminActions -notcontains $_ }).Count -eq 0)
        }
        
        # Check if delegate actions are customized
        if ($Results.AuditDelegate.Count -gt 0) {
            $delegateCustomized = -not ($Results.AuditDelegate.Count -eq $defaultDelegateActions.Count -and 
                ($Results.AuditDelegate | Where-Object { $defaultDelegateActions -notcontains $_ }).Count -eq 0)
        }
        
        # Set custom audit settings flag
        $Results.CustomAuditSettings = $ownerCustomized -or $adminCustomized -or $delegateCustomized -or $Results.AuditLogAgeLimit -ne 90
        
        # Add warnings if custom audit settings are detected
        if ($Results.CustomAuditSettings) {
            $Results.Warnings += "Mailbox has custom audit log settings that need to be reconfigured after migration"
            Write-Log -Message "Warning: Mailbox $EmailAddress has custom audit log settings" -Level "WARNING"
            
            if ($Results.AuditLogAgeLimit -ne 90) {
                Write-Log -Message "  - Custom audit log age limit: $($Results.AuditLogAgeLimit) days" -Level "WARNING"
            }
            
            if ($ownerCustomized) {
                Write-Log -Message "  - Custom owner audit actions: $($Results.AuditOwner -join ', ')" -Level "WARNING"
            }
            
            if ($adminCustomized) {
                Write-Log -Message "  - Custom admin audit actions: $($Results.AuditAdmin -join ', ')" -Level "WARNING"
            }
            
            if ($delegateCustomized) {
                Write-Log -Message "  - Custom delegate audit actions: $($Results.AuditDelegate -join ', ')" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Document these settings and reconfigure them in Exchange Online after migration" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check mailbox audit configuration: $_"
        Write-Log -Message "Warning: Failed to check mailbox audit configuration for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
