function Test-MailboxFolderStructure {
    <#
    .SYNOPSIS
        Tests a mailbox's folder structure for potential migration issues.
    
    .DESCRIPTION
        Analyzes a mailbox's folder structure to identify deeply nested folders
        and folders with large item counts that may cause problems during migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MailboxFolderStructure -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking mailbox folder structure: $EmailAddress" -Level "INFO"
        
        # Get folder statistics
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        
        # Check folder count
        $folderCount = $folderStats.Count
        $Results.FolderCount = $folderCount
        
        # Check for deep folder hierarchy (more than 10 levels might cause issues)
        $deepFolders = $folderStats | Where-Object { ($_.FolderPath.Split('/').Count - 1) -gt 10 }
        
        if ($deepFolders) {
            $Results.HasDeepFolderHierarchy = $true
            $Results.DeepFolders = $deepFolders | Select-Object Name, FolderPath, ItemsInFolder
            $Results.Warnings += "Mailbox has deeply nested folders that might cause migration issues"
            $Results.ErrorCodes += "ERR023"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has deeply nested folders:" -Level "WARNING" -ErrorCode "ERR023"
            foreach ($folder in $deepFolders | Select-Object -First 5) {
                $depth = $folder.FolderPath.Split('/').Count - 1
                Write-Log -Message "  - $($folder.FolderPath) (Depth: $depth, Items: $($folder.ItemsInFolder))" -Level "WARNING"
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
        $largeFolders = $folderStats | Where-Object { $_.ItemsInFolder -gt 5000 }
        
        if ($largeFolders) {
            $Results.HasLargeFolders = $true
            $Results.LargeFolders = $largeFolders | Select-Object Name, FolderPath, ItemsInFolder
            
            $veryLargeFolders = $largeFolders | Where-Object { $_.ItemsInFolder -gt 50000 }
            if ($veryLargeFolders) {
                $Results.Warnings += "Mailbox has folders with more than 50,000 items, which might cause migration issues"
                $Results.ErrorCodes += "ERR024"
                Write-Log -Message "Warning: Mailbox $EmailAddress has folders with extremely high item counts:" -Level "WARNING" -ErrorCode "ERR024"
            }
            else {
                Write-Log -Message "Note: Mailbox $EmailAddress has folders with high item counts:" -Level "INFO"
            }
            
            foreach ($folder in $largeFolders | Sort-Object ItemsInFolder -Descending | Select-Object -First 5) {
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
            foreach ($folder in $longNameFolders | Select-Object -First 3) {
                Write-Log -Message "  - Folder with name length $($folder.Name.Length) chars: $($folder.FolderPath)" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Consider renaming folders with extremely long names" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check mailbox folder structure: $_"
        Write-Log -Message "Warning: Failed to check mailbox folder structure for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
