function Test-ItemAgeDistribution {
    <#
    .SYNOPSIS
        Tests the age distribution of items in a mailbox.
    
    .DESCRIPTION
        Analyzes the age distribution of items in a mailbox to identify potential
        archiving needs or migration performance considerations.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "a.van.daalen@itandcare.nl"
        Test-ItemAgeDistribution -EmailAddress "a.van.daalen@itandcare.nl" -Results $results
    
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
    
    if (-not ($script:Config.CheckItemAgeDistribution)) {
        Write-Log -Message "Skipping item age distribution check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        Write-Log -Message "Analyzing item age distribution for: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "HasOldItems" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "OldestItemAge" -NotePropertyValue 0 -Force
        $Results | Add-Member -NotePropertyName "ItemAgeDistribution" -NotePropertyValue $null -Force
        
        # Get threshold for old items from config
        $ageThresholdDays = $script:Config.ItemAgeThresholdDays ?? 3650  # Default 10 years
        
        # Using Search-Mailbox is no longer recommended in Exchange Online
        # Instead, we'll use a combination of Get-MailboxFolderStatistics and Message Tracking Logs
        # to estimate item age distribution
        
        # Get folder statistics
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        
        # Get creation date of the mailbox
        $mailbox = Get-Mailbox -Identity $EmailAddress
        $creationDate = $mailbox.WhenCreated
        $mailboxAgeDays = (New-TimeSpan -Start $creationDate -End (Get-Date)).Days
        
        # Calculate age distribution estimate based on folder creation dates
        $ageDistribution = @{
            "0-1 years" = 0
            "1-3 years" = 0
            "3-5 years" = 0
            "5-10 years" = 0
            "10+ years" = 0
        }
        
        $oldestFolderDays = 0
        $oldestFolderName = ""
        
        foreach ($folder in $folderStats) {
            if ($folder.CreationTime) {
                $folderAgeDays = (New-TimeSpan -Start $folder.CreationTime -End (Get-Date)).Days
                
                # Track oldest folder
                if ($folderAgeDays -gt $oldestFolderDays) {
                    $oldestFolderDays = $folderAgeDays
                    $oldestFolderName = $folder.FolderPath
                }
                
                # Distribute items based on folder age
                # Note: This is an approximation as items could have been moved between folders
                $itemCount = $folder.ItemsInFolder
                
                if ($folderAgeDays -le 365) {
                    $ageDistribution["0-1 years"] += $itemCount
                }
                elseif ($folderAgeDays -le 1095) {
                    $ageDistribution["1-3 years"] += $itemCount
                }
                elseif ($folderAgeDays -le 1825) {
                    $ageDistribution["3-5 years"] += $itemCount
                }
                elseif ($folderAgeDays -le 3650) {
                    $ageDistribution["5-10 years"] += $itemCount
                }
                else {
                    $ageDistribution["10+ years"] += $itemCount
                }
            }
        }
        
        # Set results
        $Results.OldestItemAge = $oldestFolderDays
        $Results.ItemAgeDistribution = [PSCustomObject]$ageDistribution
        $Results.HasOldItems = $oldestFolderDays -gt $ageThresholdDays
        
        # Add warnings if old items are found
        $oldItemsCount = $ageDistribution["10+ years"]
        $veryOldItemsCount = $oldItemsCount
        
        if ($oldItemsCount -gt 0) {
            $Results.Warnings += "Mailbox contains approximately $oldItemsCount items older than 10 years"
            Write-Log -Message "Warning: Mailbox $EmailAddress has approximately $oldItemsCount items older than 10 years" -Level "WARNING"
            Write-Log -Message "  - Oldest folder: $oldestFolderName, Age: $([Math]::Round($oldestFolderDays / 365, 1)) years" -Level "WARNING"
            Write-Log -Message "Recommendation: Consider archiving old items before migration to improve performance" -Level "INFO"
        }
        
        # Log overall age distribution
        Write-Log -Message "Item age distribution for $($EmailAddress):" -Level "INFO"
        Write-Log -Message "  - 0-1 years: $($ageDistribution["0-1 years"]) items" -Level "INFO"
        Write-Log -Message "  - 1-3 years: $($ageDistribution["1-3 years"]) items" -Level "INFO"
        Write-Log -Message "  - 3-5 years: $($ageDistribution["3-5 years"]) items" -Level "INFO"
        Write-Log -Message "  - 5-10 years: $($ageDistribution["5-10 years"]) items" -Level "INFO"
        Write-Log -Message "  - 10+ years: $($ageDistribution["10+ years"]) items" -Level "INFO"
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to analyze item age distribution: $_"
        Write-Log -Message "Warning: Failed to analyze item age distribution for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
