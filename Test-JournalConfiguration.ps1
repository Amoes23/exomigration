function Test-JournalConfiguration {
    <#
    .SYNOPSIS
        Tests journaling configuration for a mailbox.
    
    .DESCRIPTION
        Checks if a mailbox is part of journaling rules that need to be preserved
        during migration to Exchange Online.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-JournalConfiguration -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking journaling configuration for: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "IsJournalRecipient" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "IsPartOfJournalRule" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "JournalRules" -NotePropertyValue @() -Force
        
        # Get journal rules and check if mailbox is a journaling recipient
        $journalRules = Get-JournalRule -ErrorAction SilentlyContinue
        $isJournalRecipient = $journalRules | Where-Object { $_.JournalEmailAddress -eq $EmailAddress }
        
        if ($isJournalRecipient) {
            $Results.IsJournalRecipient = $true
            $Results.JournalRules += ($isJournalRecipient | Select-Object Name, Scope, Enabled)
            
            $Results.Warnings += "Mailbox is configured as a journal recipient in on-premises Exchange"
            $Results.ErrorCodes += "ERR034"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress is configured as a journal recipient" -Level "WARNING" -ErrorCode "ERR034"
            foreach ($rule in $isJournalRecipient) {
                Write-Log -Message "  - Journal Rule: $($rule.Name), Scope: $($rule.Scope), Enabled: $($rule.Enabled)" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Configure journaling in Exchange Online separately after migration" -Level "INFO"
        }
        
        # Check if mailbox is part of a journal rule scope
        $mailboxInScope = $journalRules | Where-Object { 
            ($_.Scope -eq "Global") -or 
            ($_.Scope -eq "InternalMessages" -and $Results.RecipientType -ne "MailContact") -or
            ($_.Recipient -eq $EmailAddress) -or
            ($_.Recipient -like "*$EmailAddress*")
        }
        
        if ($mailboxInScope) {
            $Results.IsPartOfJournalRule = $true
            $Results.JournalRules += ($mailboxInScope | Select-Object Name, Scope, JournalEmailAddress, Enabled, Recipient)
            
            $Results.Warnings += "Mailbox is in scope for journaling rules that need to be recreated in Exchange Online"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress is in scope for journaling rules" -Level "WARNING"
            foreach ($rule in $mailboxInScope) {
                Write-Log -Message "  - Rule: $($rule.Name), Scope: $($rule.Scope), Recipient: $($rule.JournalEmailAddress)" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Document journaling rules and configure in Exchange Online" -Level "INFO"
        }
        
        if (-not $Results.IsJournalRecipient -and -not $Results.IsPartOfJournalRule) {
            Write-Log -Message "No journaling configuration found for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check journaling configuration: $_"
        Write-Log -Message "Warning: Failed to check journaling configuration for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
