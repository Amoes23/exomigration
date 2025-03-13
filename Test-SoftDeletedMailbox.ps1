function Test-SoftDeletedMailbox {
    <#
    .SYNOPSIS
        Tests for soft-deleted mailboxes that could conflict with migration.
    
    .DESCRIPTION
        Checks if the mailbox is soft-deleted or if there are soft-deleted mailboxes
        with the same identity that could cause conflicts during migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-SoftDeletedMailbox -EmailAddress "user@contoso.com" -Results $results
    
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
    
    if (-not ($script:Config.CheckForSoftDeletedItems)) {
        Write-Log -Message "Skipping soft-deleted mailbox check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        Write-Log -Message "Checking for soft-deleted mailboxes: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "IsSoftDeleted" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "HasSoftDeletedConflicts" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "SoftDeletedConflicts" -NotePropertyValue @() -Force
        
        # Get the current mailbox to extract identifying information
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        $upn = $mailbox.UserPrincipalName
        $displayName = $mailbox.DisplayName
        $alias = $mailbox.Alias
        
        # Check if the mailbox itself is soft-deleted
        # Note: This is unlikely since we just got the mailbox, but checking for completeness
        $currentSoftDeleted = Get-SoftDeletedMailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
        if ($currentSoftDeleted) {
            $Results.IsSoftDeleted = $true
            $Results.Errors += "The mailbox is currently in a soft-deleted state"
            $Results.ErrorCodes += "ERR036"
            
            Write-Log -Message "Error: Mailbox $EmailAddress is in a soft-deleted state" -Level "ERROR" -ErrorCode "ERR036"
            Write-Log -Message "  - Deletion date: $($currentSoftDeleted.WhenSoftDeleted)" -Level "ERROR"
            Write-Log -Message "Recommendation: Restore the mailbox with Restore-SoftDeletedMailbox before proceeding" -Level "INFO"
            
            return $true
        }
        
        # Check for other soft-deleted mailboxes that might conflict
        $conflicts = @()
        
        # Search by UPN
        $upnConflicts = Get-SoftDeletedMailbox -Filter "UserPrincipalName -eq '$upn'" -ErrorAction SilentlyContinue
        if ($upnConflicts) {
            foreach ($conflict in $upnConflicts) {
                $conflicts += [PSCustomObject]@{
                    Identity = $conflict.Identity
                    PrimarySmtpAddress = $conflict.PrimarySmtpAddress
                    WhenSoftDeleted = $conflict.WhenSoftDeleted
                    DisplayName = $conflict.DisplayName
                    MatchType = "UserPrincipalName"
                }
            }
        }
        
        # Search by Display Name (less precise, might find false positives)
        if ($displayName -and $displayName -ne "") {
            $displayNameConflicts = Get-SoftDeletedMailbox -Filter "DisplayName -eq '$displayName'" -ErrorAction SilentlyContinue
            if ($displayNameConflicts) {
                foreach ($conflict in $displayNameConflicts) {
                    # Skip if already found by UPN
                    if ($conflicts | Where-Object { $_.Identity -eq $conflict.Identity }) {
                        continue
                    }
                    
                    $conflicts += [PSCustomObject]@{
                        Identity = $conflict.Identity
                        PrimarySmtpAddress = $conflict.PrimarySmtpAddress
                        WhenSoftDeleted = $conflict.WhenSoftDeleted
                        DisplayName = $conflict.DisplayName
                        MatchType = "DisplayName"
                    }
                }
            }
        }
        
        # Search by Alias
        if ($alias -and $alias -ne "") {
            $aliasConflicts = Get-SoftDeletedMailbox -Filter "Alias -eq '$alias'" -ErrorAction SilentlyContinue
            if ($aliasConflicts) {
                foreach ($conflict in $aliasConflicts) {
                    # Skip if already found by UPN or DisplayName
                    if ($conflicts | Where-Object { $_.Identity -eq $conflict.Identity }) {
                        continue
                    }
                    
                    $conflicts += [PSCustomObject]@{
                        Identity = $conflict.Identity
                        PrimarySmtpAddress = $conflict.PrimarySmtpAddress
                        WhenSoftDeleted = $conflict.WhenSoftDeleted
                        DisplayName = $conflict.DisplayName
                        MatchType = "Alias"
                    }
                }
            }
        }
        
        if ($conflicts.Count -gt 0) {
            $Results.HasSoftDeletedConflicts = $true
            $Results.SoftDeletedConflicts = $conflicts
            
            $Results.Errors += "Found $($conflicts.Count) soft-deleted mailboxes that could conflict with migration"
            $Results.ErrorCodes += "ERR037"
            
            Write-Log -Message "Error: Found $($conflicts.Count) soft-deleted mailboxes that could conflict with $EmailAddress" -Level "ERROR" -ErrorCode "ERR037"
            foreach ($conflict in $conflicts) {
                Write-Log -Message "  - $($conflict.Identity) (matched by $($conflict.MatchType)), deleted on $($conflict.WhenSoftDeleted)" -Level "ERROR"
            }
            
            Write-Log -Message "Recommendation: Permanently remove these soft-deleted mailboxes with Remove-SoftDeletedMailbox" -Level "INFO"
            Write-Log -Message "Example: Remove-SoftDeletedMailbox -Identity '$($conflicts[0].Identity)' -PermanentlyDelete" -Level "INFO"
        }
        else {
            Write-Log -Message "No soft-deleted mailbox conflicts found for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for soft-deleted mailboxes: $_"
        Write-Log -Message "Warning: Failed to check for soft-deleted mailboxes for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
