function Test-ArchiveMailbox {
    <#
    .SYNOPSIS
        Tests archive mailbox configuration and size for migration readiness.
    
    .DESCRIPTION
        Analyzes the archive mailbox to ensure it meets requirements for migration
        to Exchange Online, including size limits based on license type.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-ArchiveMailbox -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking archive mailbox for: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "HasArchive" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "ArchiveSizeGB" -NotePropertyValue 0 -Force
        $Results | Add-Member -NotePropertyName "ArchiveItemCount" -NotePropertyValue 0 -Force
        $Results | Add-Member -NotePropertyName "ArchiveAutoExpandingEnabled" -NotePropertyValue $false -Force
        
        # Get mailbox to check archive status
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Check if archive is enabled
        if ($mailbox.ArchiveStatus -eq "Active") {
            $Results.HasArchive = $true
            $Results.ArchiveAutoExpandingEnabled = $mailbox.AutoExpandingArchiveEnabled -eq $true
            
            # Try to get archive statistics
            try {
                $archiveStats = Get-MailboxStatistics -Identity $EmailAddress -Archive -ErrorAction Stop
                
                if ($archiveStats) {
                    # Get archive size
                    if ($archiveStats.TotalItemSize.Value) {
                        $Results.ArchiveSizeGB = [math]::Round($archiveStats.TotalItemSize.Value.ToBytes() / 1GB, 2)
                    }
                    else {
                        # If TotalItemSize is not available, use a calculated value from item count and average size
                        $estimatedSizeBytes = $archiveStats.ItemCount * 75KB # Estimate average item size as 75KB
                        $Results.ArchiveSizeGB = [math]::Round($estimatedSizeBytes / 1GB, 2)
                    }
                    
                    # Get item count
                    $Results.ArchiveItemCount = $archiveStats.ItemCount
                    
                    Write-Log -Message "Archive mailbox size: $($Results.ArchiveSizeGB) GB with $($Results.ArchiveItemCount) items" -Level "INFO"
                }
            }
            catch {
                Write-Log -Message "Could not retrieve archive statistics: $_" -Level "WARNING"
                $Results.ArchiveSizeGB = 0
                $Results.ArchiveItemCount = 0
            }
            
            # Check against license-specific limits
            $licenseType = $Results.LicenseType ?? "Unknown"
            
            # Get archive quota based on license type
            $archiveQuotaGB = 50  # Default for lower license tiers
            
            if ($script:Config.ArchiveQuotaGB -and $script:Config.ArchiveQuotaGB.ContainsKey($licenseType)) {
                $archiveQuotaGB = $script:Config.ArchiveQuotaGB[$licenseType]
            }
            else {
                # Fallback to hardcoded values
                switch ($licenseType) {
                    "E5" { $archiveQuotaGB = 1000 }  # Unlimited archive, practically 1TB
                    "E3" { $archiveQuotaGB = 1000 }  # Unlimited archive, practically 1TB
                    "E1" { $archiveQuotaGB = 50 }    # 50GB limit for E1
                    default { $archiveQuotaGB = 50 } # Default to the lower tier
                }
            }
            
            # Check if auto-expanding archive is enabled
            $autoExpandEnabled = $Results.ArchiveAutoExpandingEnabled
            
            # If archive is approaching or exceeding limit
            if ($Results.ArchiveSizeGB -gt ($archiveQuotaGB * 0.9) -and -not $autoExpandEnabled) {
                $Results.Errors += "Archive mailbox size ($($Results.ArchiveSizeGB) GB) is approaching or exceeding license limit ($archiveQuotaGB GB) and auto-expanding archive is not enabled"
                $Results.ErrorCodes += "ERR039"
                
                Write-Log -Message "Error: Archive mailbox size ($($Results.ArchiveSizeGB) GB) is approaching license limit ($archiveQuotaGB GB)" -Level "ERROR" -ErrorCode "ERR039"
                Write-Log -Message "Recommendation: Enable auto-expanding archive or obtain a license with larger archive quota" -Level "INFO"
            }
            elseif ($Results.ArchiveSizeGB -gt ($archiveQuotaGB * 0.8) -and -not $autoExpandEnabled) {
                $Results.Warnings += "Archive mailbox size ($($Results.ArchiveSizeGB) GB) is getting close to license limit ($archiveQuotaGB GB) and auto-expanding archive is not enabled"
                
                Write-Log -Message "Warning: Archive mailbox size ($($Results.ArchiveSizeGB) GB) is approaching license limit ($archiveQuotaGB GB)" -Level "WARNING"
                Write-Log -Message "Recommendation: Consider enabling auto-expanding archive" -Level "INFO"
            }
            
            # Check if auto-expanding archive is enabled
            if ($autoExpandEnabled) {
                Write-Log -Message "Auto-expanding archive is enabled" -Level "INFO"
            }
            else {
                Write-Log -Message "Auto-expanding archive is not enabled" -Level "INFO"
            }
        }
        else {
            Write-Log -Message "Mailbox $EmailAddress does not have an archive mailbox" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check archive mailbox: $_"
        Write-Log -Message "Warning: Failed to check archive mailbox for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
