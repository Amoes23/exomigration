function Test-MailboxActivityLevel {
    <#
    .SYNOPSIS
        Tests the activity level of a mailbox for migration planning.
    
    .DESCRIPTION
        Analyzes mailbox usage patterns to estimate activity level, which can help
        with migration scheduling to minimize user disruption.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MailboxActivityLevel -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Analyzing mailbox activity level for: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "ActivityLevel" -NotePropertyValue "Low" -Force
        $Results | Add-Member -NotePropertyName "AverageDailyEmailCount" -NotePropertyValue 0 -Force
        $Results | Add-Member -NotePropertyName "LastActive" -NotePropertyValue $null -Force
        
        # Get mailbox statistics
        $stats = Get-MailboxStatistics -Identity $EmailAddress -ErrorAction Stop
        
        # If LastLogonTime is available, use it
        if ($stats.LastLogonTime) {
            $Results.LastActive = $stats.LastLogonTime
            $daysSinceLastActive = [math]::Round(((Get-Date) - $stats.LastLogonTime).TotalDays, 1)
            
            if ($daysSinceLastActive -gt 30) {
                $Results.ActivityLevel = "Inactive"
                Write-Log -Message "Mailbox $EmailAddress is inactive (no login for $daysSinceLastActive days)" -Level "INFO"
                
                return $true
            }
        }
        
        # Get message tracking logs for the past X days to estimate activity
        $activityDays = $script:Config.MaxMailboxActivityDays ?? 7  # Default to last 7 days
        $startDate = (Get-Date).AddDays(-$activityDays)
        $endDate = Get-Date
        
        # Note: This command might need to be adjusted based on environment
        # For on-premises, Get-MessageTrackingLog is available
        # For Exchange Online, this needs to be done via custom reporting
        try {
            $sentMessages = Get-MessageTrackingLog -Sender $EmailAddress -Start $startDate -End $endDate -ResultSize 5000 -ErrorAction SilentlyContinue
            $receivedMessages = Get-MessageTrackingLog -Recipients $EmailAddress -Start $startDate -End $endDate -ResultSize 5000 -ErrorAction SilentlyContinue
            
            $sentCount = if ($sentMessages) { $sentMessages.Count } else { 0 }
            $receivedCount = if ($receivedMessages) { $receivedMessages.Count } else { 0 }
            
            $totalMessages = $sentCount + $receivedCount
            $averageDailyMessages = [math]::Round($totalMessages / $activityDays, 1)
            
            $Results.AverageDailyEmailCount = $averageDailyMessages
            
            # Determine activity level based on message volume
            if ($averageDailyMessages -gt 100) {
                $Results.ActivityLevel = "Very High"
            }
            elseif ($averageDailyMessages -gt 50) {
                $Results.ActivityLevel = "High"
            }
            elseif ($averageDailyMessages -gt 20) {
                $Results.ActivityLevel = "Medium"
            }
            elseif ($averageDailyMessages -gt 5) {
                $Results.ActivityLevel = "Low"
            }
            else {
                $Results.ActivityLevel = "Very Low"
            }
            
            Write-Log -Message "Mailbox $EmailAddress activity level: $($Results.ActivityLevel) ($averageDailyMessages emails/day)" -Level "INFO"
            
            # Add recommendations based on activity level
            if ($Results.ActivityLevel -in @("High", "Very High")) {
                $Results.Warnings += "Mailbox has high activity level ($averageDailyMessages emails/day), which may require careful migration scheduling"
                Write-Log -Message "Warning: Mailbox has high activity ($averageDailyMessages emails/day), consider scheduling migration during off-hours" -Level "WARNING"
            }
        }
        catch {
            Write-Log -Message "Could not retrieve message tracking logs: $_" -Level "WARNING"
            Write-Log -Message "Using fallback method to estimate activity level" -Level "INFO"
            
            # Fallback to estimating based on item count and mailbox age
            $itemCount = $stats.ItemCount
            $creationDate = $stats.WhenCreated ?? (Get-Mailbox -Identity $EmailAddress).WhenCreated
            $mailboxAgeDays = if ($creationDate) { ((Get-Date) - $creationDate).TotalDays } else { 365 }  # Default to 1 year if unknown
            
            $avgItemsPerDay = [math]::Round($itemCount / $mailboxAgeDays, 1)
            
            # Rough estimate - each email becomes approximately 2.5 items (message + attachments + drafts)
            $estimatedEmailsPerDay = [math]::Round($avgItemsPerDay / 2.5, 1)
            $Results.AverageDailyEmailCount = $estimatedEmailsPerDay
            
            # Determine estimated activity level
            if ($estimatedEmailsPerDay -gt 40) {
                $Results.ActivityLevel = "High"
            }
            elseif ($estimatedEmailsPerDay -gt 15) {
                $Results.ActivityLevel = "Medium"
            }
            else {
                $Results.ActivityLevel = "Low"
            }
            
            Write-Log -Message "Mailbox $EmailAddress estimated activity level: $($Results.ActivityLevel) (approx. $estimatedEmailsPerDay emails/day)" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to analyze mailbox activity level: $_"
        Write-Log -Message "Warning: Failed to analyze mailbox activity level for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
