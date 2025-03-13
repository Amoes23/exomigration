function Test-MailboxConfiguration {
    <#
    .SYNOPSIS
        Tests the basic configuration of a mailbox for migration readiness.
    
    .DESCRIPTION
        Validates that a mailbox is properly configured for migration to Exchange Online,
        including checking proxy addresses, domain verification, and essential mailbox properties.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MailboxConfiguration -EmailAddress "user@contoso.com" -Results $results
    
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
        # Check if this is for a migration (mailbox doesn't exist in EXO yet)
        $isMigrationScenario = $Results.PSObject.Properties.Name -contains "OnPremisesMigration" -and $Results.OnPremisesMigration -eq $true
        
        if ($isMigrationScenario) {
            # For migration scenarios, we expect the mailbox NOT to exist in Exchange Online
            Write-Log -Message "Checking recipient conflicts in Exchange Online for migration candidate: $EmailAddress" -Level "INFO"
            
            # Check for any recipient with this address
            $recipient = Get-Recipient -Identity $EmailAddress -ErrorAction SilentlyContinue
            
            if ($recipient) {
                # If recipient exists in EXO, this is an error for migration
                $Results.Errors += "A recipient already exists in Exchange Online with this address: $($recipient.RecipientType)"
                $Results.ErrorCodes += "ERR038"
                Write-Log -Message "Error: A recipient with address $EmailAddress already exists in Exchange Online as $($recipient.RecipientType)" -Level "ERROR" -ErrorCode "ERR038"
                Write-Log -Message "Troubleshooting: Remove or rename the existing recipient before migration" -Level "INFO"
                return $false
            }
            
            # Check for mail users (contacts) with same address
            $mailUser = Get-MailUser -Identity $EmailAddress -ErrorAction SilentlyContinue
            if ($mailUser) {
                $Results.Errors += "A mail user exists in Exchange Online with this address"
                $Results.ErrorCodes += "ERR033"
                Write-Log -Message "Error: A mail user with address $EmailAddress exists in Exchange Online" -Level "ERROR" -ErrorCode "ERR033"
                Write-Log -Message "Troubleshooting: The mail user must be removed before migration" -Level "INFO"
                return $false
            }
            
            # Check for alias conflicts
            Write-Log -Message "Checking for alias conflicts: $EmailAddress" -Level "INFO"
            try {
                $alias = ($EmailAddress -split '@')[0]
                $conflictingAlias = Get-Recipient -Identity $alias -ErrorAction SilentlyContinue
                
                if ($conflictingAlias) {
                    $Results.Errors += "An alias conflict exists in Exchange Online: $alias is already used by $($conflictingAlias.PrimarySmtpAddress)"
                    $Results.ErrorCodes += "ERR038"
                    Write-Log -Message "Error: Alias conflict - $alias is already used by $($conflictingAlias.PrimarySmtpAddress)" -Level "ERROR" -ErrorCode "ERR038"
                    Write-Log -Message "Troubleshooting: The conflicting alias must be changed before migration" -Level "INFO"
                    return $false
                }
            }
            catch {
                Write-Log -Message "Warning: Failed to check for alias conflicts for $EmailAddress`: $_" -Level "WARNING"
                $Results.Warnings += "Could not check for alias conflicts: $_"
            }
            
            # Check for domain verification
            if ($script:Config.CheckDomainVerification) {
                $domain = ($EmailAddress -split '@')[1]
                $acceptedDomains = Get-AcceptedDomain
                
                if ($domain -notlike "*onmicrosoft.com") {
                    $domainVerified = $acceptedDomains | Where-Object { $_.DomainName -eq $domain -and $_.DomainType -eq "Authoritative" }
                    if (-not $domainVerified) {
                        $Results.Errors += "Domain not verified in Exchange Online: $domain"
                        $Results.ErrorCodes += "ERR015"
                        Write-Log -Message "Error: Domain not verified: $domain for $EmailAddress" -Level "ERROR" -ErrorCode "ERR015"
                        Write-Log -Message "Troubleshooting: Verify domain ownership in Microsoft 365 admin center" -Level "INFO"
                    }
                }
            }
            
            # For migration scenarios, absence of a mailbox is expected and good
            $Results.ExistingMailbox = $false
            return $true
        }
        else {
            # Standard Exchange Online mailbox validation (non-migration scenario)
            # Get Mailbox info from Exchange Online
            $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
            $Results.ExistingMailbox = $true
            $Results.DisplayName = $mailbox.DisplayName
            $Results.RecipientType = $mailbox.RecipientType
            $Results.RecipientTypeDetails = $mailbox.RecipientTypeDetails
            $Results.MailboxEnabled = $true
            $Results.HiddenFromAddressList = $mailbox.HiddenFromAddressListsEnabled
            $Results.LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
            $Results.RetentionHoldEnabled = $mailbox.RetentionHoldEnabled
            $Results.ExchangeGuid = $mailbox.ExchangeGuid.ToString()
            $Results.HasLegacyExchangeDN = -not [string]::IsNullOrEmpty($mailbox.LegacyExchangeDN)
            $Results.ArchiveStatus = if ($mailbox.ArchiveStatus) { $mailbox.ArchiveStatus } else { "NotEnabled" }
            $Results.ForwardingEnabled = $mailbox.DeliverToMailboxAndForward
            $Results.ForwardingAddress = $mailbox.ForwardingAddress
            $Results.SAMAccountName = $mailbox.SamAccountName
            $Results.Alias = $mailbox.Alias
            
            # Check SAM account name and alias validity
            $Results.SAMAccountNameValid = -not [string]::IsNullOrEmpty($mailbox.SamAccountName)
            $Results.AliasValid = -not [string]::IsNullOrEmpty($mailbox.Alias)
            
            # Check for both litigation and retention hold enabled
            if ($mailbox.LitigationHoldEnabled -and $mailbox.RetentionHoldEnabled) {
                $Results.Warnings += "Both Litigation Hold and Retention Hold are enabled"
                Write-Log -Message "Warning: Both Litigation Hold and Retention Hold are enabled for $EmailAddress" -Level "WARNING"
            }
            
            # Check primary SMTP address
            $primarySMTP = ($mailbox.EmailAddresses | Where-Object { $_ -clike "SMTP:*" }).Substring(5)
            
            # Check UPN versus Primary SMTP
            if ($mailbox.UserPrincipalName -eq $primarySMTP) {
                $Results.UPNMatchesPrimarySMTP = $true
            }
            else {
                if ($script:Config.RequireUPNMatchSMTP) {
                    $Results.Warnings += "UPN does not match primary SMTP address"
                    Write-Log -Message "Warning: UPN ($($mailbox.UserPrincipalName)) does not match primary SMTP address ($primarySMTP) for $EmailAddress" -Level "WARNING"
                    Write-Log -Message "Troubleshooting: Consider updating the UPN to match the primary SMTP address" -Level "INFO"
                }
            }
            
            $Results.UPN = $mailbox.UserPrincipalName
            
            # Check for onmicrosoft.com address
            $onMicrosoftAddresses = $mailbox.EmailAddresses | Where-Object { $_ -like "*onmicrosoft.com*" }
            $targetDeliveryDomain = $script:Config.TargetDeliveryDomain
            
            if ($onMicrosoftAddresses.Count -gt 0) {
                $Results.HasOnMicrosoftAddress = $true
                
                if ($targetDeliveryDomain) {
                    $hasTargetDomain = $onMicrosoftAddresses | Where-Object { $_ -like "*$targetDeliveryDomain*" }
                    if ($hasTargetDomain) {
                        $Results.HasRequiredOnMicrosoftAddress = $true
                    }
                    else {
                        $Results.Warnings += "Mailbox has onmicrosoft.com address but not with the required domain: $targetDeliveryDomain"
                        Write-Log -Message "Warning: Mailbox has onmicrosoft.com address but not with the required domain: $targetDeliveryDomain" -Level "WARNING"
                    }
                }
            }
            else {
                $Results.Errors += "No onmicrosoft.com email address found"
                $Results.ErrorCodes += "ERR014"
                Write-Log -Message "Error: No onmicrosoft.com email address found for $EmailAddress" -Level "ERROR" -ErrorCode "ERR014"
                Write-Log -Message "Troubleshooting: Add an email alias with your tenant's onmicrosoft.com domain" -Level "INFO"
            }
            
            # Check proxy address formatting
            $primaryCount = ($mailbox.EmailAddresses | Where-Object { $_ -clike "SMTP:*" }).Count
            if ($primaryCount -ne 1) {
                $Results.ProxyAddressesFormatting = $false
                $Results.Errors += "Invalid number of primary SMTP addresses: $primaryCount"
                $Results.ErrorCodes += "ERR011"
                Write-Log -Message "Error: Invalid number of primary SMTP addresses ($primaryCount) for $EmailAddress" -Level "ERROR" -ErrorCode "ERR011"
                Write-Log -Message "Troubleshooting: Ensure exactly one primary SMTP address is set (capitalized SMTP:)" -Level "INFO"
            }
            
            # Check domain verification for all SMTP addresses if enabled in config
            if ($script:Config.CheckDomainVerification) {
                $allDomainsVerified = $true
                $unverifiedDomains = @()
                $acceptedDomains = Get-AcceptedDomain
                
                foreach ($address in $mailbox.EmailAddresses) {
                    if ($address -like "smtp:*" -or $address -like "SMTP:*") {
                        $domain = ($address -split "@")[1]
                        if ($domain -notlike "*onmicrosoft.com") {
                            $domainVerified = $acceptedDomains | Where-Object { $_.DomainName -eq $domain -and $_.DomainType -eq "Authoritative" }
                            if (-not $domainVerified) {
                                $allDomainsVerified = $false
                                $unverifiedDomains += $domain
                                
                                if (-not $Results.HasUnverifiedDomains) {
                                    $Results.HasUnverifiedDomains = $true
                                    $Results.Errors += "Domain not verified: $domain"
                                    $Results.ErrorCodes += "ERR015"
                                    Write-Log -Message "Error: Domain not verified: $domain for $EmailAddress" -Level "ERROR" -ErrorCode "ERR015"
                                    Write-Log -Message "Troubleshooting: Verify domain ownership in Microsoft 365 admin center" -Level "INFO"
                                }
                            }
                        }
                    }
                }
                
                $Results.AllDomainsVerified = $allDomainsVerified
                $Results.UnverifiedDomains = $unverifiedDomains
            }
            else {
                $Results.AllDomainsVerified = $true
            }
            
            return $true
        }
    }
    catch {
        # If this is a migration scenario, not finding the mailbox in Exchange Online is expected
        if ($isMigrationScenario) {
            if ($_.Exception.Message -like "*couldn't be found*") {
                Write-Log -Message "No mailbox found in Exchange Online for $EmailAddress (expected for migration)" -Level "INFO"
                $Results.ExistingMailbox = $false
                return $true
            }
        }
        
        # Otherwise, report as an error
        $Results.Errors += "Failed to get mailbox: $_"
        Write-Log -Message "Error: Failed to check mailbox configuration for $EmailAddress`: $_" -Level "ERROR"
        return $false
    }
}