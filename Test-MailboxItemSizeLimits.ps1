function Test-MailboxItemSizeLimits {
    <#
    .SYNOPSIS
        Tests a mailbox for items exceeding size limits that may affect migration.
    
    .DESCRIPTION
        Checks for large items in a mailbox that may exceed size limits for migration.
        Identifies folders containing oversized items and reports details to help
        with remediation before migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .PARAMETER MaxItemSizeMB
        Maximum allowed item size in MB. Default is 150 MB.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MailboxItemSizeLimits -EmailAddress "user@contoso.com" -Results $results
    
    .EXAMPLE
        Test-MailboxItemSizeLimits -EmailAddress "user@contoso.com" -Results $results -MaxItemSizeMB 100
    
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
        [int]$MaxItemSizeMB = 0
    )
    
    try {
        # Use config value if not specified directly
        if ($MaxItemSizeMB -eq 0 -and $script:Config.MaxItemSizeMB) {
            $MaxItemSizeMB = $script:Config.MaxItemSizeMB
        }
        elseif ($MaxItemSizeMB -eq 0) {
            $MaxItemSizeMB = 150  # Default value if not in config
        }
        
        Write-Log -Message "Checking for items exceeding $MaxItemSizeMB MB in mailbox: $EmailAddress" -Level "INFO"
        
        # Get the mailbox
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Use Get-MailboxFolderStatistics to check for large items
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        $largeItems = @()
        $foundLargeItems = $false
        
        foreach ($folder in $folderStats) {
            if ($folder.MaxItemSize -and $folder.MaxItemSize -ne "0 B") {
                # Extract numeric part from size string (like "157.3 MB" -> 157.3)
                $sizeValue = $folder.MaxItemSize
                
                # Convert MaxItemSize to numeric (MB)
                if ($sizeValue -like "*GB*") {
                    $sizeInMB = [double]($sizeValue -replace '[^\d\.]', '') * 1024
                }
                elseif ($sizeValue -like "*MB*") {
                    $sizeInMB = [double]($sizeValue -replace '[^\d\.]', '')
                }
                elseif ($sizeValue -like "*KB*") {
                    $sizeInMB = [double]($sizeValue -replace '[^\d\.]', '') / 1024
                }
                elseif ($sizeValue -like "*B*") {
                    $sizeInMB = [double]($sizeValue -replace '[^\d\.]', '') / (1024 * 1024)
                }
                else {
                    # Assume bytes if no unit is specified
                    $sizeInMB = [double]($sizeValue -replace '[^\d\.]', '') / (1024 * 1024)
                }
                
                if ($sizeInMB -ge $MaxItemSizeMB) {
                    $foundLargeItems = $true
                    $largeItems += [PSCustomObject]@{
                        FolderName = $folder.Name
                        FolderPath = $folder.FolderPath
                        MaxItemSize = $folder.MaxItemSize
                        ItemCount = $folder.ItemsInFolder
                        SizeInMB = $sizeInMB
                    }
                }
            }
        }
        
        if ($foundLargeItems) {
            $Results.HasLargeItems = $true
            $Results.LargeItemsDetails = $largeItems
            
            $largestItem = $largeItems | Sort-Object -Property SizeInMB -Descending | Select-Object -First 1
            
            $Results.Warnings += "Mailbox contains items larger than $MaxItemSizeMB MB (largest: $($largestItem.MaxItemSize)), which may cause migration issues"
            $Results.ErrorCodes += "ERR019"
            
            Write-Log -Message "Warning: Found items exceeding $MaxItemSizeMB MB in mailbox $EmailAddress" -Level "WARNING" -ErrorCode "ERR019"
            Write-Log -Message "Large items found in the following folders:" -Level "WARNING"
            
            foreach ($item in ($largeItems | Sort-Object -Property SizeInMB -Descending | Select-Object -First 5)) {
                Write-Log -Message "  - Folder: $($item.FolderPath), Max item size: $($item.MaxItemSize)" -Level "WARNING"
            }
            
            if ($largeItems.Count -gt 5) {
                Write-Log -Message "  - ... and $($largeItems.Count - 5) more folders with large items" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Review and remove/archive large items before migration" -Level "INFO"
            Write-Log -Message "Consider using -BadItemLimit parameter if large items cannot be removed" -Level "INFO"
        }
        else {
            $Results.HasLargeItems = $false
            Write-Log -Message "No items larger than $MaxItemSizeMB MB found in mailbox $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for large items: $_"
        Write-Log -Message "Warning: Failed to check for large items in mailbox $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
