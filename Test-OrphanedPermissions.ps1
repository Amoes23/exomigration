function Test-OrphanedPermissions {
    <#
    .SYNOPSIS
        Tests for orphaned permissions on a mailbox.
    
    .DESCRIPTION
        Checks if a mailbox has permissions assigned to users who no longer exist
        or are disabled, which may cause issues during migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-OrphanedPermissions -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking for orphaned permissions: $EmailAddress" -Level "INFO"
        
        # Get mailbox permissions
        $permissions = Get-MailboxPermission -Identity $EmailAddress | Where-Object { 
            $_.AccessRights -contains "FullAccess" -and 
            $_.User -notlike "NT AUTHORITY\*" -and 
            $_.User -ne "Default" -and
            $_.Deny -eq $false
        }
        
        $orphanedPermissions = @()
        
        foreach ($permission in $permissions) {
            $user = $permission.User.ToString()
            
            # Skip if the user is a security principal (not a string representation)
            if ($permission.User -isnot [string]) {
                continue
            }
            
            # Check if the user exists
            try {
                $recipient = Get-Recipient -Identity $user -ErrorAction Stop
                # User exists, check if it's disabled
                if ($recipient.RecipientTypeDetails -eq "DisabledUser") {
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = $permission.AccessRights -join ", "
                        Issue = "User account is disabled"
                    }
                }
            }
            catch [Microsoft.Exchange.Management.Tasks.ManagementObjectNotFoundException] {
                # User not found in directory
                $orphanedPermissions += [PSCustomObject]@{
                    User = $user
                    AccessRights = $permission.AccessRights -join ", "
                    Issue = "User not found in directory"
                }
            }
            catch [System.Management.Automation.RemoteException] {
                # Handle remote exceptions (common with Exchange Online)
                if ($_.Exception.Message -like "*not found*") {
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = $permission.AccessRights -join ", "
                        Issue = "User not found in directory (remote error)"
                    }
                }
                else {
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = $permission.AccessRights -join ", "
                        Issue = "Error checking user: $($_.Exception.Message)"
                    }
                }
            }
            catch {
                # User might not exist anymore
                $orphanedPermissions += [PSCustomObject]@{
                    User = $user
                    AccessRights = $permission.AccessRights -join ", "
                    Issue = "User not found in directory"
                }
            }
        }
        
        # Check Send on Behalf permissions for orphaned users
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        if ($mailbox.GrantSendOnBehalfTo) {
            foreach ($user in $mailbox.GrantSendOnBehalfTo) {
                try {
                    $recipient = Get-Recipient -Identity $user -ErrorAction Stop
                    # User exists, check if it's disabled
                    if ($recipient.RecipientTypeDetails -eq "DisabledUser") {
                        $orphanedPermissions += [PSCustomObject]@{
                            User = $user
                            AccessRights = "SendOnBehalf"
                            Issue = "User account is disabled"
                        }
                    }
                }
                catch {
                    # User might not exist anymore
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = "SendOnBehalf"
                        Issue = "User not found in directory"
                    }
                }
            }
        }
        
        # Check Send As permissions for orphaned users
        $sendAsPermissions = Get-RecipientPermission -Identity $EmailAddress | Where-Object {
            $_.AccessRights -contains "SendAs" -and 
            $_.Trustee -notlike "NT AUTHORITY\*" -and 
            $_.Trustee -ne "Default"
        }
        
        foreach ($permission in $sendAsPermissions) {
            $user = $permission.Trustee.ToString()
            
            try {
                $recipient = Get-Recipient -Identity $user -ErrorAction Stop
                # User exists, check if it's disabled
                if ($recipient.RecipientTypeDetails -eq "DisabledUser") {
                    $orphanedPermissions += [PSCustomObject]@{
                        User = $user
                        AccessRights = "SendAs"
                        Issue = "User account is disabled"
                    }
                }
            }
            catch {
                # User might not exist anymore
                $orphanedPermissions += [PSCustomObject]@{
                    User = $user
                    AccessRights = "SendAs"
                    Issue = "User not found in directory"
                }
            }
        }
        
        if ($orphanedPermissions.Count -gt 0) {
            $Results.HasOrphanedPermissions = $true
            $Results.OrphanedPermissions = $orphanedPermissions
            $Results.Warnings += "Mailbox has orphaned permissions that may not migrate correctly"
            $Results.ErrorCodes += "ERR022"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has orphaned permissions:" -Level "WARNING" -ErrorCode "ERR022"
            foreach ($perm in $orphanedPermissions) {
                Write-Log -Message "  - User: $($perm.User), Rights: $($perm.AccessRights), Issue: $($perm.Issue)" -Level "WARNING"
            }
            Write-Log -Message "Recommendation: Clean up orphaned permissions before migration" -Level "INFO"
        }
        else {
            $Results.HasOrphanedPermissions = $false
            Write-Log -Message "No orphaned permissions found for mailbox $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for orphaned permissions: $_"
        Write-Log -Message "Warning: Failed to check for orphaned permissions for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
