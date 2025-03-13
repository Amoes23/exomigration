function Test-OnPremRecoverableItemsSize {
    <#
    .SYNOPSIS
        Tests the size of the Recoverable Items folder in an on-premises mailbox.
    
    .DESCRIPTION
        Analyzes the size of the Recoverable Items folder of an on-premises mailbox
        to identify potential issues during migration, such as exceeding quota limits
        or causing migration delays.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-OnPremRecoverableItemsSize -EmailAddress "user@contoso.com" -Results $results
    
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
    
    if (-not ($script:Config.CheckRecoverableItemsSize)) {
        Write-Log -Message "Skipping recoverable items size check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        Write-Log -Message "Checking recoverable items folder size for on-premises mailbox: $EmailAddress" -Level "INFO"
        
        # Get folder statistics for the mailbox
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress -FolderScope RecoverableItems -ErrorAction Stop
        
        if ($folderStats) {
            # Initialize properties
            $Results | Add-Member -NotePropertyName "RecoverableItemsFolderSizeGB" -NotePropertyValue 0 -Force
            $Results | Add-Member -NotePropertyName "RecoverableItemsCount" -NotePropertyValue 0 -Force
            $Results | Add-Member -NotePropertyName "RecoverableItemsFolderDetails" -NotePropertyValue @() -Force
            
            # Calculate total size of all recoverable items folders
            $totalSizeBytes = 0
            $totalItems = 0
            $folderDetails = @()
            
            foreach ($folder in $folderStats) {
                # Skip folders that aren't part of Recoverable Items
                if ($folder.FolderPath -notlike "/Recoverable Items*") {
                    continue
                }
                
                # Get size in bytes 
                if ($folder.FolderAndSubfolderSize) {
                    # Parse size - format varies between versions
                    if ($folder.FolderAndSubfolderSize -match "([0-9,.]+)\s+(B|KB|MB|GB)") {
                        $sizeValue = [double]($Matches[1] -replace ',', '')
                        $sizeUnit = $Matches[2]
                        
                        # Convert to bytes
                        $sizeBytes = switch ($sizeUnit) {
                            "B" { $sizeValue }
                            "KB" { $sizeValue * 1KB }
                            "MB" { $sizeValue * 1MB }
                            "GB" { $sizeValue * 1GB }
                            default { 0 }
                        }
                        
                        $totalSizeBytes += $sizeBytes
                    }
                    
                    $totalItems += $folder.ItemsInFolderAndSubfolders
                    
                    # Add folder details
                    $folderDetails += [PSCustomObject]@{
                        FolderPath = $folder.FolderPath
                        ItemCount = $folder.ItemsInFolderAndSubfolders
                        FolderSize = $folder.FolderAndSubfolderSize
                        SizeBytes = $sizeBytes
                    }
                }
            }
            
            # Set properties
            $Results.RecoverableItemsFolderSizeGB = [math]::Round($totalSizeBytes / 1GB, 2)
            $Results.RecoverableItemsCount = $totalItems
            $Results.RecoverableItemsFolderDetails = $folderDetails
            
            # Get threshold from config
            $sizeThresholdGB = $script:Config.RecoverableItemsSizeThresholdGB ?? 15
            
            # Check if size exceeds threshold
            if ($Results.RecoverableItemsFolderSizeGB -gt $sizeThresholdGB) {
                $Results.Warnings += "Recoverable Items folder is very large ($($Results.RecoverableItemsFolderSizeGB) GB), which may delay migration"
                Write-Log -Message "Warning: Mailbox $EmailAddress has a large Recoverable Items folder ($($Results.RecoverableItemsFolderSizeGB) GB)" -Level "WARNING"
                Write-Log -Message "Recommendation: Consider clearing litigation hold or cleaning up Recoverable Items before migration" -Level "INFO"
                Write-Log -Message "Note: Large Recoverable Items folders will take longer to migrate and may require specific migration batch settings" -Level "INFO"
            }
            else {
                Write-Log -Message "Recoverable Items folder size for $EmailAddress is $($Results.RecoverableItemsFolderSizeGB) GB ($totalItems items)" -Level "INFO"
            }
            
            # Check for specific subfolder issues
            $deletionsFolder = $folderDetails | Where-Object { $_.FolderPath -like "*Deletions*" }
            $versionsFolder = $folderDetails | Where-Object { $_.FolderPath -like "*Versions*" }
            $purgesFolder = $folderDetails | Where-Object { $_.FolderPath -like "*Purges*" }
            
            if ($deletionsFolder -and ($deletionsFolder.SizeBytes / 1GB) -gt 5) {
                $Results.Warnings += "Deletions folder is large ($([math]::Round($deletionsFolder.SizeBytes / 1GB, 2)) GB), consider clearing before migration"
                Write-Log -Message "Warning: Deletions folder is large ($([math]::Round($deletionsFolder.SizeBytes / 1GB, 2)) GB) for $EmailAddress" -Level "WARNING"
            }
            
            if ($purgesFolder -and ($purgesFolder.SizeBytes / 1GB) -gt 3) {
                $Results.Warnings += "Purges folder is large ($([math]::Round($purgesFolder.SizeBytes / 1GB, 2)) GB), which may indicate retention policy issues"
                Write-Log -Message "Warning: Purges folder is large ($([math]::Round($purgesFolder.SizeBytes / 1GB, 2)) GB) for $EmailAddress" -Level "WARNING"
            }
            
            # Check mailbox for retention and litigation holds
            $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
            if ($mailbox -and ($mailbox.LitigationHoldEnabled -or $mailbox.RetentionHoldEnabled)) {
                $Results.Warnings += "Mailbox has " + 
                    $(if ($mailbox.LitigationHoldEnabled) {"litigation hold"} else {""}) + 
                    $(if ($mailbox.LitigationHoldEnabled -and $mailbox.RetentionHoldEnabled) {" and "} else {""}) +
                    $(if ($mailbox.RetentionHoldEnabled) {"retention hold"} else {""}) + 
                    " enabled, which may cause large Recoverable Items folders"
                
                Write-Log -Message "Warning: Mailbox $EmailAddress has holds enabled that may impact Recoverable Items size" -Level "WARNING"
                Write-Log -Message "Recommendation: Review hold settings and determine if they need to be preserved in Exchange Online" -Level "INFO"
            }
        }
        else {
            Write-Log -Message "No Recoverable Items folder statistics found for $EmailAddress" -Level "INFO"
            
            # Initialize properties with default values
            $Results | Add-Member -NotePropertyName "RecoverableItemsFolderSizeGB" -NotePropertyValue 0 -Force
            $Results | Add-Member -NotePropertyName "RecoverableItemsCount" -NotePropertyValue 0 -Force
            $Results | Add-Member -NotePropertyName "RecoverableItemsFolderDetails" -NotePropertyValue @() -Force
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check Recoverable Items folder size: $_"
        Write-Log -Message "Warning: Failed to check Recoverable Items folder size for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}