function Test-SpecialMailboxTypes {
    <#
    .SYNOPSIS
        Tests if a mailbox is a special type that requires specific migration handling.
    
    .DESCRIPTION
        Checks if the mailbox is a special type such as shared, resource, discovery, 
        or other special-purpose mailbox that requires specific handling during migration.
        Provides guidance on the appropriate migration approach for each type.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-SpecialMailboxTypes -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking for special mailbox type: $EmailAddress" -Level "INFO"
        
        # Ensure properties exist
        $Results | Add-Member -NotePropertyName "IsSpecialMailbox" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "SpecialMailboxType" -NotePropertyValue $null -Force
        $Results | Add-Member -NotePropertyName "SpecialMailboxGuidance" -NotePropertyValue $null -Force
        
        # Get mailbox to check type
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Define special mailbox types and their guidance
        $specialTypes = @{
            "SharedMailbox" = @{
                Guidance = "Ensure all users with access to this shared mailbox are migrated together. Permissions will need to be verified post-migration."
                Warning = "Shared mailbox detected - requires specific permission handling during migration"
            }
            "RoomMailbox" = @{
                Guidance = "Room mailboxes require calendar processing settings to be verified post-migration. Check booking policies and delegate access."
                Warning = "Room mailbox detected - requires calendar processing verification post-migration"
            }
            "EquipmentMailbox" = @{
                Guidance = "Equipment mailboxes require calendar processing settings to be verified post-migration. Check booking policies and delegate access."
                Warning = "Equipment mailbox detected - requires calendar processing verification post-migration"
            }
            "DiscoveryMailbox" = @{
                Guidance = "Discovery mailboxes should be handled separately from regular user migrations. Consider creating a new discovery mailbox in Exchange Online instead of migrating."
                Warning = "Discovery mailbox detected - not recommended for standard migration"
            }
            "MonitoringMailbox" = @{
                Guidance = "Monitoring mailboxes typically should not be migrated. Consider creating new monitoring configurations in Exchange Online."
                Warning = "Monitoring mailbox detected - not recommended for standard migration"
            }
            "LinkedMailbox" = @{
                Guidance = "Linked mailboxes must be converted to regular mailboxes before migration. This requires specific preparation steps."
                Warning = "Linked mailbox detected - requires conversion before migration"
            }
            "TeamMailbox" = @{
                Guidance = "Team mailboxes should be replaced with Microsoft 365 Groups instead of being directly migrated."
                Warning = "Team mailbox detected - consider using Microsoft 365 Groups instead"
            }
            "ArbitrationMailbox" = @{
                Guidance = "Arbitration mailboxes should not be migrated through the standard migration process. They require special handling."
                Warning = "Arbitration mailbox detected - requires special migration process"
            }
            "AuditLogMailbox" = @{
                Guidance = "Audit log mailboxes should not be migrated. They are automatically created in Exchange Online."
                Warning = "Audit log mailbox detected - should not be migrated"
            }
            "JournalMailbox" = @{
                Guidance = "Journal mailboxes should not be migrated. Configure journaling in Exchange Online separately."
                Warning = "Journal mailbox detected - should not be migrated"
            }
            "PublicFolderMailbox" = @{
                Guidance = "Public folder mailboxes require a separate migration process using the public folder migration cmdlets."
                Warning = "Public folder mailbox detected - requires separate migration process"
            }
            "GroupMailbox" = @{
                Guidance = "Group mailboxes should be migrated using the group migration process, not the mailbox migration process."
                Warning = "Group mailbox detected - requires separate group migration process"
            }
        }
        
        # Check if this is a special mailbox based on RecipientTypeDetails
        $recipientTypeDetails = $mailbox.RecipientTypeDetails.ToString()
        
        if ($specialTypes.ContainsKey($recipientTypeDetails)) {
            $Results.IsSpecialMailbox = $true
            $Results.SpecialMailboxType = $recipientTypeDetails
            $Results.SpecialMailboxGuidance = $specialTypes[$recipientTypeDetails].Guidance
            
            $Results.Warnings += $specialTypes[$recipientTypeDetails].Warning
            $Results.ErrorCodes += "ERR026"
            
            Write-Log -Message "Special mailbox type detected for $EmailAddress`: $recipientTypeDetails" -Level "WARNING" -ErrorCode "ERR026"
            Write-Log -Message "  Guidance: $($specialTypes[$recipientTypeDetails].Guidance)" -Level "INFO"
        }
        else {
            # Check if it's a non-user mailbox based on other properties
            if ($mailbox.IsShared -eq $true) {
                $Results.IsSpecialMailbox = $true
                $Results.SpecialMailboxType = "SharedMailbox"
                $Results.SpecialMailboxGuidance = $specialTypes["SharedMailbox"].Guidance
                
                $Results.Warnings += $specialTypes["SharedMailbox"].Warning
                
                Write-Log -Message "Shared mailbox detected for $EmailAddress based on IsShared property" -Level "WARNING"
                Write-Log -Message "  Guidance: $($specialTypes["SharedMailbox"].Guidance)" -Level "INFO"
            }
            elseif ($mailbox.IsResource -eq $true) {
                $type = if ($mailbox.ResourceType -eq "Room") { "RoomMailbox" } else { "EquipmentMailbox" }
                
                $Results.IsSpecialMailbox = $true
                $Results.SpecialMailboxType = $type
                $Results.SpecialMailboxGuidance = $specialTypes[$type].Guidance
                
                $Results.Warnings += $specialTypes[$type].Warning
                
                Write-Log -Message "$type detected for $EmailAddress based on IsResource property" -Level "WARNING"
                Write-Log -Message "  Guidance: $($specialTypes[$type].Guidance)" -Level "INFO"
            }
            else {
                Write-Log -Message "Standard user mailbox detected for $EmailAddress" -Level "INFO"
            }
        }
        
        # Additional check for specific attributes that indicate special purpose mailboxes
        if ($mailbox.CustomAttribute1 -eq "DiscoverySearchMailbox") {
            $Results.IsSpecialMailbox = $true
            $Results.SpecialMailboxType = "DiscoveryMailbox"
            $Results.SpecialMailboxGuidance = $specialTypes["DiscoveryMailbox"].Guidance
            
            $Results.Warnings += $specialTypes["DiscoveryMailbox"].Warning
            
            Write-Log -Message "Discovery mailbox detected for $EmailAddress based on CustomAttribute1" -Level "WARNING"
            Write-Log -Message "  Guidance: $($specialTypes["DiscoveryMailbox"].Guidance)" -Level "INFO"
        }
        
        # Check if mailbox is part of journal configuration
        try {
            $journalRules = Get-JournalRule -ErrorAction SilentlyContinue
            $isJournalRecipient = $journalRules | Where-Object { $_.RecipientEmailAddress -eq $EmailAddress }
            
            if ($isJournalRecipient) {
                $Results.IsSpecialMailbox = $true
                $Results.SpecialMailboxType = "JournalMailbox"
                $Results.SpecialMailboxGuidance = $specialTypes["JournalMailbox"].Guidance
                
                $Results.Warnings += $specialTypes["JournalMailbox"].Warning
                
                Write-Log -Message "Journal mailbox detected for $EmailAddress" -Level "WARNING"
                Write-Log -Message "  Guidance: $($specialTypes["JournalMailbox"].Guidance)" -Level "INFO"
            }
        }
        catch {
            Write-Log -Message "Could not check journal rules: $_" -Level "DEBUG"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for special mailbox type: $_"
        Write-Log -Message "Warning: Failed to check for special mailbox type for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
