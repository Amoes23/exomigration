function Test-OnPremMailboxFolderStructure {
    <#
    .SYNOPSIS
        Tests an on-premises mailbox's folder structure for potential migration issues.
    
    .DESCRIPTION
        Analyzes an on-premises mailbox's folder structure to identify deeply nested folders
        and folders with large item counts that may cause problems during migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-OnPremMailboxFolderStructure -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking on-premises mailbox folder structure: $EmailAddress" -Level "INFO"
        
        # Get folder statistics
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        
        # Check folder count
        $folderCount = $folderStats.Count
        $Results.FolderCount = $folderCount
        
        # Check for folder count limits
        $maxFolderLimit = $script:Config.MaxFolderLimit ?? 10000
        if ($folderCount -gt $maxFolderLimit) {
            $Results.Warnings += "Mailbox has an excessive number of folders ($folderCount), which might cause migration issues"
            Write-Log -Message "Warning: Mailbox $EmailAddress has an excessive number of folders: $folderCount" -Level "WARNING"
            Write-Log -Message "Recommendation: Consider cleaning up unnecessary folders before migration" -Level "INFO"
        }
        
        # Check for deep folder hierarchy (more than 10 levels might cause issues)
        $deepFolders = @()
        $maxDepth = $script:Config.MaxTotalFolderDepth ?? 10

        foreach ($folder in $folderStats) {
            # Calculate folder depth by counting path separators
            $depth = ($folder.FolderPath.Split('/').Count - 1)
            
            if ($depth -gt $maxDepth) {
                $deepFolders += [PSCustomObject]@{
                    Name = $folder.Name
                    FolderPath = $folder.FolderPath
                    ItemsInFolder = $folder.ItemsInFolder
                    Depth = $depth
                }
            }
        }
        
        if ($deepFolders.Count -gt 0) {
            $Results.HasDeepFolderHierarchy = $true
            $Results.DeepFolders = $deepFolders
            $Results.Warnings += "Mailbox has deeply nested folders that might cause migration issues"
            $Results.ErrorCodes += "ERR023"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has deeply nested folders:" -Level "WARNING" -ErrorCode "ERR023"
            foreach ($folder in ($deepFolders | Select-Object -First 5)) {
                Write-Log -Message "  - $($folder.FolderPath) (Depth: $($folder.Depth), Items: $($folder.ItemsInFolder))" -Level "WARNING"
            }
            
            if ($deepFolders.Count -gt 5) {
                Write-Log -Message "  - ... and $($deepFolders.Count - 5) more deeply nested folders" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Simplify folder structure before migration" -Level "INFO"
        }
        else {
            $Results.HasDeepFolderHierarchy = $false
            Write-Log -Message "No deeply nested folders found in mailbox $EmailAddress" -Level "INFO"
        }
        
        # Check for large folders (more than 5000 items might cause performance issues)
        $largeFolders = @()
        $maxItemsPerFolder = $script:Config.MaxItemsPerFolderLimit ?? 100000
        $warningItemsPerFolder = $maxItemsPerFolder / 2
        
        foreach ($folder in $folderStats) {
            if ($folder.ItemsInFolder -gt $warningItemsPerFolder) {
                $largeFolders += [PSCustomObject]@{
                    Name = $folder.Name
                    FolderPath = $folder.FolderPath
                    ItemsInFolder = $folder.ItemsInFolder
                }
            }
        }
        
        if ($largeFolders.Count -gt 0) {
            $Results.HasLargeFolders = $true
            $Results.LargeFolders = $largeFolders
            
            $veryLargeFolders = $largeFolders | Where-Object { $_.ItemsInFolder -gt $maxItemsPerFolder }
            if ($veryLargeFolders) {
                $Results.Warnings += "Mailbox has folders with more than $($maxItemsPerFolder) items, which might cause migration issues"
                $Results.ErrorCodes += "ERR024"
                Write-Log -Message "Warning: Mailbox $EmailAddress has folders with extremely high item counts:" -Level "WARNING" -ErrorCode "ERR024"
            }
            else {
                Write-Log -Message "Note: Mailbox $EmailAddress has folders with high item counts:" -Level "INFO"
            }
            
            foreach ($folder in ($largeFolders | Sort-Object ItemsInFolder -Descending | Select-Object -First 5)) {
                Write-Log -Message "  - $($folder.FolderPath): $($folder.ItemsInFolder) items" -Level "INFO"
            }
            
            if ($largeFolders.Count -gt 5) {
                Write-Log -Message "  - ... and $($largeFolders.Count - 5) more large folders" -Level "INFO"
            }
            
            if ($veryLargeFolders) {
                Write-Log -Message "Recommendation: Consider archiving items from large folders to improve migration performance" -Level "INFO"
            }
        }
        else {
            $Results.HasLargeFolders = $false
            Write-Log -Message "No folders with high item counts found in mailbox $EmailAddress" -Level "INFO"
        }
        
        # Get total item count
        $totalItems = ($folderStats | Measure-Object -Property ItemsInFolder -Sum).Sum
        $Results.TotalItemCount = $totalItems
        
        if ($totalItems -gt 100000) {
            $Results.Warnings += "Mailbox has a very high total item count ($totalItems items), which might slow down migration"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a very high total item count: $totalItems items" -Level "WARNING"
            Write-Log -Message "Recommendation: Consider archiving older items before migration to improve performance" -Level "INFO"
        }
        else {
            Write-Log -Message "Mailbox $EmailAddress has $totalItems items in total" -Level "INFO"
        }
        
        # Check folder name length
        $longNameFolders = $folderStats | Where-Object { $_.Name.Length -gt 100 }
        if ($longNameFolders) {
            $Results.Warnings += "Mailbox has folders with very long names that might cause issues during migration"
            Write-Log -Message "Warning: Mailbox $EmailAddress has folders with very long names:" -Level "WARNING"
            foreach ($folder in ($longNameFolders | Select-Object -First 3)) {
                Write-Log -Message "  - Folder with name length $($folder.Name.Length) chars: $($folder.FolderPath)" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Consider renaming folders with extremely long names" -Level "INFO"
        }
        
        # Check for problematic folder names (special characters)
        $problemFolders = $folderStats | Where-Object { 
            $_.Name -match '[<>:"\/\\|?*]' -or
            $_.Name -match '^\s+' -or  # Leading spaces
            $_.Name -match '\s+      # Trailing spaces
        }
        
        if ($problemFolders) {
            $Results.Warnings += "Mailbox has folders with problematic names (special characters or leading/trailing spaces)"
            Write-Log -Message "Warning: Mailbox $EmailAddress has folders with problematic characters:" -Level "WARNING"
            
            foreach ($folder in ($problemFolders | Select-Object -First 3)) {
                Write-Log -Message "  - Problematic folder name: '$($folder.Name)': $($folder.FolderPath)" -Level "WARNING"
            }
            
            if ($problemFolders.Count -gt 3) {
                Write-Log -Message "  - ... and $($problemFolders.Count - 3) more problematic folder names" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Rename folders with special characters before migration" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check mailbox folder structure: $_"
        Write-Log -Message "Warning: Failed to check mailbox folder structure for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}