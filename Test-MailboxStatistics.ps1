# Update to Test-MailboxStatistics.ps1
# Add secondary threshold for "large but within quota" mailboxes

function Test-MailboxStatistics {
    <#
    .SYNOPSIS
        Tests the statistics of a mailbox for migration readiness.
    
    .DESCRIPTION
        Retrieves and analyzes mailbox statistics including size, item count,
        and other metrics that may impact migration performance.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MailboxStatistics -EmailAddress "user@contoso.com" -Results $results
    
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
        # Get mailbox statistics
        $stats = Get-MailboxStatistics -Identity $EmailAddress -ErrorAction Stop
        
        # Size calculations
        if ($stats.TotalItemSize.Value) {
            $Results.MailboxSizeGB = [math]::Round($stats.TotalItemSize.Value.ToBytes() / 1GB, 2)
        }
        else {
            # If TotalItemSize is not available, use a calculated value from item count and average size
            $estimatedSizeBytes = $stats.ItemCount * 75KB # Estimate average item size as 75KB
            $Results.MailboxSizeGB = [math]::Round($estimatedSizeBytes / 1GB, 2)
            
            $Results.Warnings += "Mailbox size is estimated based on item count"
            Write-Log -Message "Warning: Mailbox size for $EmailAddress is estimated based on item count" -Level "WARNING"
        }
        
        # Determine if mailbox exceeds size threshold
        $maxMailboxSize = $script:Config.MaxMailboxSizeGB ?? 50
        
        # Check if near or exceeding quota based on license type
        $nearQuotaPercentage = $script:Config.NearQuotaPercentageThreshold ?? 80
        $nearQuotaThreshold = $maxMailboxSize * ($nearQuotaPercentage / 100)
        
        # Add license-aware size checking
        $licenseType = $Results.LicenseType ?? "Unknown"
        $mailboxSizeGB = $Results.MailboxSizeGB

        # License-specific size thresholds
        switch ($licenseType) {
            "E5" {
                $absoluteQuota = 100
                $archiveQuota = 1000  # Effectively unlimited with auto-expanding archives
            }
            "E3" {
                $absoluteQuota = 100
                $archiveQuota = 1000  # Effectively unlimited with auto-expanding archives
            }
            "E1" {
                $absoluteQuota = 50
                $archiveQuota = 50
            }
            default {
                $absoluteQuota = $maxMailboxSize
                $archiveQuota = 50
            }
        }
        
        # Add property for license-specific quota
        $Results | Add-Member -NotePropertyName "LicenseSpecificQuotaGB" -NotePropertyValue $absoluteQuota -Force

        # Check if exceeding license quota
        if ($mailboxSizeGB -gt $absoluteQuota) {
            $Results.MailboxSizeWarning = $true
            $Results.Warnings += "Mailbox size ($mailboxSizeGB GB) exceeds $licenseType license quota ($absoluteQuota GB)"
            Write-Log -Message "Warning: Mailbox size ($mailboxSizeGB GB) exceeds $licenseType license quota ($absoluteQuota GB) for $EmailAddress" -Level "WARNING"
            Write-Log -Message "Recommendation: Consider archiving or data cleanup before migration" -Level "INFO"
        }
        # NEW: Check if approaching license quota (e.g., 80% of quota)
        elseif ($mailboxSizeGB -gt ($absoluteQuota * ($nearQuotaPercentage / 100))) {
            $Results | Add-Member -NotePropertyName "MailboxNearQuota" -NotePropertyValue $true -Force
            $Results.Warnings += "Mailbox size ($mailboxSizeGB GB) is approaching $licenseType license quota ($absoluteQuota GB) - currently at $([math]::Round(($mailboxSizeGB / $absoluteQuota) * 100))%"
            Write-Log -Message "Warning: Mailbox size ($mailboxSizeGB GB) is approaching $licenseType license quota ($absoluteQuota GB) for $EmailAddress" -Level "WARNING"
            Write-Log -Message "Performance Impact: Large mailboxes may experience longer migration times" -Level "INFO"
        }
        else {
            $Results | Add-Member -NotePropertyName "MailboxNearQuota" -NotePropertyValue $false -Force
        }
        
        # Get item count
        $Results.TotalItemCount = $stats.ItemCount
        
        # Warn about excessive item count
        if ($Results.TotalItemCount -gt 100000) {
            $Results.Warnings += "Mailbox has a very high item count ($($Results.TotalItemCount) items), which might slow down migration"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a very high item count: $($Results.TotalItemCount) items" -Level "WARNING"
            Write-Log -Message "Recommendation: Consider archiving older items before migration to improve performance" -Level "INFO"
        }
        
        # Check last logon time
        if ($stats.LastLogonTime) {
            $Results | Add-Member -NotePropertyName "LastLogonTime" -NotePropertyValue $stats.LastLogonTime -Force
            $daysSinceLastLogon = (New-TimeSpan -Start $stats.LastLogonTime -End (Get-Date)).Days
            
            $inactivityThreshold = $script:Config.InactivityThresholdDays ?? 30
            if ($daysSinceLastLogon -gt $inactivityThreshold) {
                $Results | Add-Member -NotePropertyName "IsInactive" -NotePropertyValue $true -Force
                $Results | Add-Member -NotePropertyName "InactiveDays" -NotePropertyValue $daysSinceLastLogon -Force
                
                $Results.Warnings += "Mailbox inactive for $daysSinceLastLogon days (threshold: $inactivityThreshold days)"
                Write-Log -Message "Warning: Mailbox $EmailAddress inactive for $daysSinceLastLogon days" -Level "WARNING"
            }
            else {
                $Results | Add-Member -NotePropertyName "IsInactive" -NotePropertyValue $false -Force
                $Results | Add-Member -NotePropertyName "InactiveDays" -NotePropertyValue $daysSinceLastLogon -Force
            }
        }
        else {
            $Results | Add-Member -NotePropertyName "LastLogonTime" -NotePropertyValue $null -Force
            $Results | Add-Member -NotePropertyName "IsInactive" -NotePropertyValue $true -Force
            $Results | Add-Member -NotePropertyName "InactiveDays" -NotePropertyValue 999 -Force
            
            $Results.Warnings += "Mailbox has never been logged into"
            Write-Log -Message "Warning: Mailbox $EmailAddress has never been logged into" -Level "WARNING"
        }
        
        # Check deleted items
        if ($stats.DeletedItemCount -gt 10000) {
            $Results.Warnings += "Mailbox has a large number of deleted items ($($stats.DeletedItemCount)), which might impact migration"
            Write-Log -Message "Warning: Mailbox $EmailAddress has $($stats.DeletedItemCount) deleted items" -Level "WARNING"
            Write-Log -Message "Recommendation: Consider emptying the Deleted Items folder before migration" -Level "INFO"
        }
        
        # Determine corruption risk based on size and item count
        $Results | Add-Member -NotePropertyName "PotentialCorruptionRisk" -NotePropertyValue "Low" -Force
        
        if ($Results.TotalItemCount -gt 100000 || $Results.MailboxSizeGB -gt 50) {
            $Results.PotentialCorruptionRisk = "High"
            # Recommend BadItemLimit as approximately 0.1% of total items, capped at 100
            $recommendedLimit = [Math]::Min(100, [Math]::Ceiling($Results.TotalItemCount * 0.001))
            $Results | Add-Member -NotePropertyName "RecommendedBadItemLimit" -NotePropertyValue $recommendedLimit -Force
            
            $Results.Warnings += "Large mailbox with high potential for corrupt items. Recommended BadItemLimit: $recommendedLimit"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a high risk of corrupt items due to size ($($Results.TotalItemCount) items)" -Level "WARNING"
            Write-Log -Message "Recommendation: BadItemLimit: $recommendedLimit for $EmailAddress" -Level "INFO"
        }
        elseif ($Results.TotalItemCount -gt 50000 || $Results.MailboxSizeGB -gt 25) {
            $Results.PotentialCorruptionRisk = "Medium"
            # Recommend BadItemLimit as approximately 0.05% of total items, capped at 50
            $recommendedLimit = [Math]::Min(50, [Math]::Ceiling($Results.TotalItemCount * 0.0005))
            $Results | Add-Member -NotePropertyName "RecommendedBadItemLimit" -NotePropertyValue $recommendedLimit -Force
            
            $Results.Warnings += "Medium-sized mailbox with potential for corrupt items. Recommended BadItemLimit: $recommendedLimit"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a medium risk of corrupt items ($($Results.TotalItemCount) items)" -Level "WARNING"
            Write-Log -Message "Recommendation: BadItemLimit: $recommendedLimit for $EmailAddress" -Level "INFO"
        }
        else {
            $Results | Add-Member -NotePropertyName "RecommendedBadItemLimit" -NotePropertyValue 10 -Force
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to get mailbox statistics: $_"
        Write-Log -Message "Warning: Failed to get mailbox statistics for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
