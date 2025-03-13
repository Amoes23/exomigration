function Test-MailboxPermissions {
    <#
    .SYNOPSIS
        Tests mailbox permissions for migration readiness.
    
    .DESCRIPTION
        Checks mailbox permissions including Full Access, Send As, and Send on Behalf
        permissions that may need to be preserved during migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MailboxPermissions -EmailAddress "user@contoso.com" -Results $results
    
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
        # Check for Full Access permissions
        $fullAccessPermissions = Get-MailboxPermission -Identity $EmailAddress | Where-Object { 
            $_.AccessRights -contains "FullAccess" -and 
            $_.User -notlike "NT AUTHORITY\*" -and 
            $_.User -ne "Default" -and
            $_.Deny -eq $false
        }
        
        if ($fullAccessPermissions) {
            $Results.FullAccessDelegates = $fullAccessPermissions | ForEach-Object { $_.User.ToString() }
            
            if ($fullAccessPermissions.Count -gt 5) {
                $Results.Warnings += "Mailbox has $($fullAccessPermissions.Count) Full Access delegates which may require special attention"
                Write-Log -Message "Warning: Mailbox $EmailAddress has $($fullAccessPermissions.Count) Full Access delegates" -Level "WARNING"
            }
            else {
                Write-Log -Message "Info: Full Access permissions found for $EmailAddress : $($Results.FullAccessDelegates -join ', ')" -Level "INFO"
            }
        }
        
        # Check Send on Behalf permissions
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
        if ($mailbox -and $mailbox.GrantSendOnBehalfTo -and $mailbox.GrantSendOnBehalfTo.Count -gt 0) {
            $Results.SendOnBehalfDelegates = $mailbox.GrantSendOnBehalfTo
            
            if ($mailbox.GrantSendOnBehalfTo.Count -gt 5) {
                $Results.Warnings += "Mailbox has $($mailbox.GrantSendOnBehalfTo.Count) Send on Behalf delegates which may require special attention"
                Write-Log -Message "Warning: Mailbox $EmailAddress has $($mailbox.GrantSendOnBehalfTo.Count) Send on Behalf delegates" -Level "WARNING"
            }
            else {
                Write-Log -Message "Info: Send on Behalf permissions found for $EmailAddress : $($Results.SendOnBehalfDelegates -join ', ')" -Level "INFO"
            }
        }
        
        # Check for Send As permissions
        $sendAsPermissions = Get-RecipientPermission -Identity $EmailAddress | Where-Object { 
            $_.AccessRights -contains "SendAs" -and 
            $_.Trustee -notlike "NT AUTHORITY\*" -and 
            $_.Trustee -ne "Default"
        }
        
        if ($sendAsPermissions) {
            $Results.HasSendAsPermissions = $true
            $Results.SendAsPermissions = $sendAsPermissions | ForEach-Object { $_.Trustee.ToString() }
            
            if ($sendAsPermissions.Count -gt 5) {
                $Results.Warnings += "Mailbox has $($sendAsPermissions.Count) Send As permissions which may require special attention"
                Write-Log -Message "Warning: Mailbox $EmailAddress has $($sendAsPermissions.Count) Send As permissions" -Level "WARNING" 
            }
            else {
                Write-Log -Message "Info: SendAs permissions found for $EmailAddress : $($Results.SendAsPermissions -join ', ')" -Level "INFO"
            }
        }
        
        # Check if the mailbox is a delegate for other mailboxes
        try {
            $delegateFor = @()
            
            # Check for mailboxes where this user has Full Access
            $userHasFullAccessTo = Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission -User $EmailAddress -ErrorAction SilentlyContinue | 
                Where-Object { $_.AccessRights -contains "FullAccess" -and $_.Deny -eq $false }
            
            if ($userHasFullAccessTo) {
                foreach ($permission in $userHasFullAccessTo) {
                    $delegateFor += [PSCustomObject]@{
                        Mailbox = $permission.Identity
                        PermissionType = "FullAccess"
                    }
                }
            }
            
            # Check for mailboxes where this user has Send As rights
            $userHasSendAsTo = Get-Mailbox -ResultSize Unlimited | Get-RecipientPermission -Trustee $EmailAddress -ErrorAction SilentlyContinue |
                Where-Object { $_.AccessRights -contains "SendAs" }
            
            if ($userHasSendAsTo) {
                foreach ($permission in $userHasSendAsTo) {
                    $delegateFor += [PSCustomObject]@{
                        Mailbox = $permission.Identity
                        PermissionType = "SendAs"
                    }
                }
            }
            
            # Check for mailboxes where this user has Send on Behalf rights
            $userSendOnBehalfTo = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.GrantSendOnBehalfTo -contains $EmailAddress }
            
            if ($userSendOnBehalfTo) {
                foreach ($mbx in $userSendOnBehalfTo) {
                    $delegateFor += [PSCustomObject]@{
                        Mailbox = $mbx.PrimarySmtpAddress
                        PermissionType = "SendOnBehalf"
                    }
                }
            }
            
            # Add delegate information to results if found
            if ($delegateFor.Count -gt 0) {
                $Results | Add-Member -NotePropertyName "DelegateForMailboxes" -NotePropertyValue $delegateFor -Force
                
                $Results.Warnings += "User is a delegate for $($delegateFor.Count) other mailboxes"
                Write-Log -Message "Warning: User $EmailAddress is a delegate for $($delegateFor.Count) other mailboxes" -Level "WARNING"
                Write-Log -Message "Recommendation: Ensure these mailboxes are migrated together or permissions are preserved" -Level "INFO"
            }
        }
        catch {
            # Delegate check isn't critical, so just log a warning
            Write-Log -Message "Warning: Failed to check if user is a delegate for other mailboxes: $_" -Level "WARNING"
        }
        
        # Check for folder-level permissions
        try {
            # Check if the mailbox has given folder permissions to others
            $folderPermissions = Get-MailboxFolderPermission -Identity "$($EmailAddress):\Inbox" -ErrorAction SilentlyContinue |
                Where-Object { $_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" }
            
            if ($folderPermissions -and $folderPermissions.Count -gt 0) {
                $Results | Add-Member -NotePropertyName "HasFolderPermissions" -NotePropertyValue $true -Force
                $Results | Add-Member -NotePropertyName "FolderPermissionsCount" -NotePropertyValue $folderPermissions.Count -Force
                
                $Results.Warnings += "Mailbox has custom folder permissions that may require special handling"
                Write-Log -Message "Warning: Mailbox $EmailAddress has custom folder permissions that may require special handling" -Level "WARNING"
            }
        }
        catch {
            # Folder permission check isn't critical, so just log a warning
            Write-Log -Message "Warning: Failed to check mailbox folder permissions: $_" -Level "WARNING"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check mailbox permissions: $_"
        Write-Log -Message "Warning: Failed to check mailbox permissions for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}