function Test-FolderNameConflicts {
    <#
    .SYNOPSIS
        Tests for folder name conflicts that could affect migration.
    
    .DESCRIPTION
        Checks for problematic folder names such as those with special characters,
        excessive length, or duplicates that could cause migration issues.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-FolderNameConflicts -EmailAddress "user@contoso.com" -Results $results
    
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
    
    if (-not ($script:Config.CheckFolderNameConflicts)) {
        Write-Log -Message "Skipping folder name conflicts check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        Write-Log -Message "Checking for folder name conflicts: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "HasFolderNameConflicts" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "FolderNameConflicts" -NotePropertyValue @() -Force
        
        # Get folder statistics
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        
        # Check for problematic folder names
        $conflicts = @()
        $folderNames = @{}
        $invalidChars = '<>:"/\|?*'
        
        foreach ($folder in $folderStats) {
            $folderName = $folder.Name
            $folderPath = $folder.FolderPath
            
            # Check for very long folder names (>128 chars)
            if ($folderName.Length -gt 128) {
                $conflicts += [PSCustomObject]@{
                    FolderPath = $folderPath
                    Issue = "Very long folder name ($($folderName.Length) characters)"
                    Type = "LongName"
                }
            }
            
            # Check for special characters
            $hasInvalidChars = $false
            foreach ($char in $invalidChars.ToCharArray()) {
                if ($folderName.Contains($char)) {
                    $hasInvalidChars = $true
                    break
                }
            }
            
            if ($hasInvalidChars) {
                $conflicts += [PSCustomObject]@{
                    FolderPath = $folderPath
                    Issue = "Contains special characters that may cause migration issues"
                    Type = "SpecialCharacters"
                }
            }
            
            # Check for folder name collisions at the same level
            $parentPath = Split-Path -Path $folderPath
            $key = "$parentPath|$folderName"
            
            if ($folderNames.ContainsKey($key)) {
                $conflicts += [PSCustomObject]@{
                    FolderPath = $folderPath
                    Issue = "Possible folder name collision with similar name at same level"
                    Type = "NameCollision"
                }
            }
            else {
                $folderNames[$key] = $folderPath
            }
            
            # Check for leading/trailing spaces
            if ($folderName -match '^\s|\s$') {
                $conflicts += [PSCustomObject]@{
                    FolderPath = $folderPath
                    Issue = "Folder name has leading or trailing spaces"
                    Type = "WhitespaceIssue"
                }
            }
            
            # Check for path length (>255 chars is problematic)
            if ($folderPath.Length -gt 255) {
                $conflicts += [PSCustomObject]@{
                    FolderPath = $folderPath
                    Issue = "Path length exceeds 255 characters ($($folderPath.Length))"
                    Type = "PathTooLong"
                }
            }
        }
        
        if ($conflicts.Count -gt 0) {
            $Results.HasFolderNameConflicts = $true
            $Results.FolderNameConflicts = $conflicts
            
            $Results.Warnings += "Mailbox has $($conflicts.Count) folders with naming issues that may cause migration problems"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has $($conflicts.Count) folders with name issues" -Level "WARNING"
            
            $longNameFolders = ($conflicts | Where-Object { $_.Type -eq "LongName" }).Count
            $specialCharFolders = ($conflicts | Where-Object { $_.Type -eq "SpecialCharacters" }).Count
            $collisionFolders = ($conflicts | Where-Object { $_.Type -eq "NameCollision" }).Count
            $whitespaceFolders = ($conflicts | Where-Object { $_.Type -eq "WhitespaceIssue" }).Count
            $longPathFolders = ($conflicts | Where-Object { $_.Type -eq "PathTooLong" }).Count
            
            if ($longNameFolders -gt 0) {
                Write-Log -Message "  - $longNameFolders folders with names >128 characters" -Level "WARNING"
            }
            
            if ($specialCharFolders -gt 0) {
                Write-Log -Message "  - $specialCharFolders folders with special characters" -Level "WARNING"
            }
            
            if ($collisionFolders -gt 0) {
                Write-Log -Message "  - $collisionFolders potential folder name collisions" -Level "WARNING"
            }
            
            if ($whitespaceFolders -gt 0) {
                Write-Log -Message "  - $whitespaceFolders folders with leading/trailing spaces" -Level "WARNING"
            }
            
            if ($longPathFolders -gt 0) {
                Write-Log -Message "  - $longPathFolders folders with path length >255 characters" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Review and rename problematic folders before migration" -Level "INFO"
        }
        else {
            Write-Log -Message "No folder name conflicts found for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for folder name conflicts: $_"
        Write-Log -Message "Warning: Failed to check for folder name conflicts for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
