function Test-CloudMailboxPlaceholder {
    <#
    .SYNOPSIS
        Tests for existing cloud mailbox placeholders that could conflict with migration.
    
    .DESCRIPTION
        Checks if a placeholder mailbox already exists in Exchange Online with the same
        identity, which could cause conflicts during migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-CloudMailboxPlaceholder -EmailAddress "user@contoso.com" -Results $results
    
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
    
    if (-not ($script:Config.CheckCloudMailboxPlaceholders)) {
        Write-Log -Message "Skipping cloud mailbox placeholder check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        Write-Log -Message "Checking for cloud mailbox placeholders: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "HasCloudPlaceholder" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "CloudPlaceholderDetails" -NotePropertyValue $null -Force
        
        # Get the on-premises mailbox to compare against
        try {
            $sourceMailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
            $userPrincipalName = $sourceMailbox.UserPrincipalName
            $exchangeGuid = $sourceMailbox.ExchangeGuid
            $mailNickname = $sourceMailbox.Alias
        }
        catch [Microsoft.Exchange.Management.Tasks.ManagementObjectNotFoundException] {
            Write-Log -Message "Source mailbox not found: $EmailAddress" -Level "ERROR"
            $Results.Errors += "Source mailbox not found - check if the mailbox exists in the source environment"
            return $false
        }
        catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
            Write-Log -Message "Insufficient permissions to access source mailbox: $EmailAddress" -Level "ERROR"
            $Results.Errors += "Insufficient permissions to access source mailbox information"
            return $false
        }
        catch {
            Write-Log -Message "Error retrieving source mailbox: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
            $Results.Errors += "Failed to retrieve source mailbox: $($_.Exception.Message)"
            return $false
        }
        
        # Check for recipient with same UPN
        try {
            $recipientByUPN = Get-Recipient -Identity $userPrincipalName -ErrorAction SilentlyContinue -RecipientTypeDetails RemovedMailbox,SoftDeletedMailbox,MailUser,User,Guest
        }
        catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
            Write-Log -Message "Insufficient permissions to check for UPN recipient: $userPrincipalName" -Level "WARNING"
            $Results.Warnings += "Insufficient permissions to check for UPN recipient"
            $recipientByUPN = $null
        }
        catch {
            Write-Log -Message "Error checking for UPN recipient: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
            $Results.Warnings += "Error checking for UPN recipient: $($_.Exception.Message)"
            $recipientByUPN = $null
        }
        
        # Check for recipient with same ExchangeGuid
        try {
            $recipientByGuid = Get-Recipient -Filter "ExchangeGuid -eq '$exchangeGuid'" -ErrorAction SilentlyContinue
        }
        catch [System.Management.Automation.ParameterBindingException] {
            Write-Log -Message "Invalid ExchangeGuid format for filter: $exchangeGuid" -Level "WARNING"
            $Results.Warnings += "Invalid ExchangeGuid format: $exchangeGuid"
            $recipientByGuid = $null
        }
        catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
            Write-Log -Message "Insufficient permissions to check for ExchangeGuid recipient" -Level "WARNING"
            $Results.Warnings += "Insufficient permissions to check for ExchangeGuid recipient"
            $recipientByGuid = $null
        }
        catch {
            Write-Log -Message "Error checking for ExchangeGuid recipient: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
            $Results.Warnings += "Error checking for ExchangeGuid recipient: $($_.Exception.Message)"
            $recipientByGuid = $null
        }
        
        # Check for recipient with same alias
        try {
            $recipientByAlias = Get-Recipient -Identity $mailNickname -ErrorAction SilentlyContinue -RecipientTypeDetails RemovedMailbox,SoftDeletedMailbox,MailUser,User,Guest
        }
        catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
            Write-Log -Message "Insufficient permissions to check for alias recipient: $mailNickname" -Level "WARNING"
            $Results.Warnings += "Insufficient permissions to check for alias recipient"
            $recipientByAlias = $null
        }
        catch {
            Write-Log -Message "Error checking for alias recipient: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
            $Results.Warnings += "Error checking for alias recipient: $($_.Exception.Message)"
            $recipientByAlias = $null
        }
        
        # Check RecipientTypeDetails to identify soft-deleted or mail-enabled users
        # that could cause conflicts
        $placeholder = $null
        $placeholderType = "None"
        
        if ($recipientByGuid) {
            $placeholder = $recipientByGuid
            $placeholderType = "ExchangeGuid"
        }
        elseif ($recipientByUPN -and $recipientByUPN.PrimarySmtpAddress -ne $sourceMailbox.PrimarySmtpAddress) {
            $placeholder = $recipientByUPN
            $placeholderType = "UserPrincipalName"
        }
        elseif ($recipientByAlias -and $recipientByAlias.PrimarySmtpAddress -ne $sourceMailbox.PrimarySmtpAddress) {
            $placeholder = $recipientByAlias
            $placeholderType = "Alias"
        }
        
        if ($placeholder) {
            $Results.HasCloudPlaceholder = $true
            $Results.CloudPlaceholderDetails = [PSCustomObject]@{
                Identity = $placeholder.Identity
                PrimarySmtpAddress = $placeholder.PrimarySmtpAddress
                RecipientTypeDetails = $placeholder.RecipientTypeDetails
                MatchType = $placeholderType
                WhenCreated = $placeholder.WhenCreated
                IsSoftDeleted = $placeholder.RecipientTypeDetails -in @("SoftDeletedMailbox", "RemovedMailbox")
                IsMailUser = $placeholder.RecipientTypeDetails -eq "MailUser"
            }
            
            # Add appropriate warnings or errors based on placeholder type
            $placeholderTypeName = $placeholder.RecipientTypeDetails.ToString()
            
            if ($placeholderTypeName -in @("SoftDeletedMailbox", "RemovedMailbox")) {
                $Results.Errors += "A soft-deleted mailbox with matching $placeholderType exists in Exchange Online"
                $Results.ErrorCodes += "ERR032"
                
                Write-Log -Message "Error: Soft-deleted mailbox found for $EmailAddress (matching $placeholderType)" -Level "ERROR" -ErrorCode "ERR032"
                Write-Log -Message "  - Details: $placeholderTypeName created on $($placeholder.WhenCreated)" -Level "ERROR"
                Write-Log -Message "Recommendation: Restore or permanently remove the soft-deleted mailbox" -Level "INFO"
                Write-Log -Message "Command: Remove-MailUser -Identity '$($placeholder.Identity)' -Permanent `$true" -Level "INFO"
            }
            elseif ($placeholderTypeName -eq "MailUser") {
                $Results.Errors += "A mail user with matching $placeholderType exists in Exchange Online"
                $Results.ErrorCodes += "ERR033"
                
                Write-Log -Message "Error: Mail user found for $EmailAddress (matching $placeholderType)" -Level "ERROR" -ErrorCode "ERR033"
                Write-Log -Message "  - Details: $placeholderTypeName created on $($placeholder.WhenCreated)" -Level "ERROR"
                Write-Log -Message "Recommendation: Convert the mail user or remove it before migration" -Level "INFO"
            }
            else {
                $Results.Warnings += "A recipient with matching $placeholderType exists in Exchange Online"
                
                Write-Log -Message "Warning: Recipient found for $EmailAddress (matching $placeholderType)" -Level "WARNING"
                Write-Log -Message "  - Details: $placeholderTypeName created on $($placeholder.WhenCreated)" -Level "WARNING"
                Write-Log -Message "Recommendation: Resolve the conflict by renaming or removing the existing recipient" -Level "INFO"
            }
        }
        else {
            Write-Log -Message "No cloud mailbox placeholder found for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Unexpected error in Test-CloudMailboxPlaceholder: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
        if ($_.ScriptStackTrace) {
            Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
        }
        $Results.Warnings += "Failed to check for cloud mailbox placeholders: $_"
        return $false
    }
}
