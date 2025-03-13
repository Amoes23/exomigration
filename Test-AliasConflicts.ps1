function Test-AliasConflicts {
    <#
    .SYNOPSIS
        Tests for alias conflicts that could prevent migration.
    
    .DESCRIPTION
        Checks if the mailbox's alias or any of its email addresses conflict with
        existing objects in Exchange Online, which could prevent successful migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-AliasConflicts -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking for alias conflicts: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "HasAliasConflicts" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "AliasConflicts" -NotePropertyValue @() -Force
        
        # Get mailbox to check aliases
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Check primary alias (mailbox nickname)
        $alias = $mailbox.Alias
        $conflicts = @()
        
        # Check for alias conflicts
        $aliasConflict = Get-Recipient -Filter "Alias -eq '$alias' -and PrimarySmtpAddress -ne '$EmailAddress'" -ErrorAction SilentlyContinue
        if ($aliasConflict) {
            foreach ($conflict in $aliasConflict) {
                $conflicts += [PSCustomObject]@{
                    ConflictType = "Alias"
                    ConflictValue = $alias
                    ConflictingObject = $conflict.PrimarySmtpAddress
                    RecipientType = $conflict.RecipientType
                }
            }
        }
        
        # Extract all email addresses from the mailbox
        $emailAddresses = @()
        foreach ($address in $mailbox.EmailAddresses) {
            if ($address -like "smtp:*") {
                $emailAddresses += $address.Substring(5)  # Remove "smtp:" prefix
            }
        }
        
        # Check for email address conflicts
        foreach ($address in $emailAddresses) {
            $addressConflict = Get-Recipient -Filter "EmailAddresses -like '*$address*' -and PrimarySmtpAddress -ne '$EmailAddress'" -ErrorAction SilentlyContinue
            if ($addressConflict) {
                foreach ($conflict in $addressConflict) {
                    $conflicts += [PSCustomObject]@{
                        ConflictType = "EmailAddress"
                        ConflictValue = $address
                        ConflictingObject = $conflict.PrimarySmtpAddress
                        RecipientType = $conflict.RecipientType
                    }
                }
            }
        }
        
        if ($conflicts.Count -gt 0) {
            $Results.HasAliasConflicts = $true
            $Results.AliasConflicts = $conflicts
            
            $Results.Errors += "Found $($conflicts.Count) alias or email address conflicts that may prevent migration"
            $Results.ErrorCodes += "ERR038"
            
            Write-Log -Message "Error: Found $($conflicts.Count) alias or email address conflicts for $EmailAddress" -Level "ERROR" -ErrorCode "ERR038"
            foreach ($conflict in $conflicts) {
                Write-Log -Message "  - $($conflict.ConflictType) conflict: $($conflict.ConflictValue) with $($conflict.ConflictingObject) ($($conflict.RecipientType))" -Level "ERROR"
            }
            
            Write-Log -Message "Recommendation: Resolve conflicts by changing aliases or removing conflicting addresses" -Level "INFO"
        }
        else {
            Write-Log -Message "No alias or email address conflicts found for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for alias conflicts: $_"
        Write-Log -Message "Warning: Failed to check for alias conflicts for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
