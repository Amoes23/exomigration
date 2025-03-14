function Test-RecoverableItemsSize {
    <#
    .SYNOPSIS
        Tests the size of the Recoverable Items folder in a mailbox.
    
    .DESCRIPTION
        Analyzes the size of the Recoverable Items folder to identify potential issues
        during migration, such as exceeding quota limits or causing migration delays.
        Works with both on-premises Exchange and Exchange Online mailboxes.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .PARAMETER OnPremises
        When specified, treats the mailbox as an on-premises mailbox.
        Otherwise, assumes Exchange Online mailbox.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-RecoverableItemsSize -EmailAddress "user@contoso.com" -Results $results
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-RecoverableItemsSize -EmailAddress "user@contoso.com" -Results $results -OnPremises
    
    .OUTPUTS
        [bool] Returns $true if the test was successful (even if issues were found), $false if the test failed.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results,
        
        [Parameter(Mandatory = $false)]
        [switch]$OnPremises
    )
    
    if (-not ($script:Config.CheckRecoverableItemsSize)) {
        Write-Log -Message "Skipping recoverable items size check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        $envType = if ($OnPremises) { "on-premises" } else { "Exchange Online" }
        Write-Log -Message "Checking $envType recoverable items folder size for: $EmailAddress" -Level "INFO"
        
        # Get folder statistics for the mailbox (works the same in both environments)
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
                
                # Get size in bytes - with on-premises vs cloud format handling
                if ($OnPremises -and $folder.FolderAndSubfolderSize) {
                    # Parse size from on-premises format
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
                        $totalItems += $folder.ItemsInFolderAndSubfolders
                    }
                }
                elseif (-not $OnPremises -and $folder.FolderSize -ne "0 B") {
                    # Parse size from cloud format
                    if ($folder.FolderSize -match "([0-9,.]+)\s+(B|KB|MB|GB)") {
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
                        $totalItems += $folder.ItemsInFolder
                    }
                }
                
                # Create common folder details object
                $itemCount = if ($OnPremises) { $folder.ItemsInFolderAndSubfolders } else { $folder.ItemsInFolder }
                $sizeProperty = if ($OnPremises) { $folder.FolderAndSubfolderSize } else { $folder.FolderSize }
                
                $folderDetails += [PSCustomObject]@{
                    FolderPath = $folder.FolderPath
                    ItemCount = $itemCount
                    FolderSize = $sizeProperty
                    SizeBytes = $sizeBytes
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
                $Results.ErrorCodes += "ERR062"
                
                Write-Log -Message "Warning: Mailbox $EmailAddress has a large Recoverable Items folder ($($Results.RecoverableItemsFolderSizeGB) GB)" -Level "WARNING" -ErrorCode "ERR062"
                Write-Log -Message "Recommendation: Consider clearing litigation hold or using New-ComplianceSearch to export data before migration" -Level "INFO"
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
                Write-Log -Message "Recommendation: Review hold settings and determine if they need to be preserved" -Level "INFO"
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
