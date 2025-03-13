function Test-FolderPermissionDepth {
    <#
    .SYNOPSIS
        Tests the depth and complexity of folder permissions in a mailbox.
    
    .DESCRIPTION
        Analyzes folder permissions to identify deeply nested or complex permission
        structures that might not migrate correctly to Exchange Online.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-FolderPermissionDepth -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Analyzing folder permission depth for: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "HasDeepFolderPermissions" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "FolderPermissionsCount" -NotePropertyValue 0 -Force
        $Results | Add-Member -NotePropertyName "MaxFolderPermissionDepth" -NotePropertyValue 0 -Force
        $Results | Add-Member -NotePropertyName "ComplexPermissionFolders" -NotePropertyValue @() -Force
        
        # Get the mailbox
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Get folder statistics
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        
        # Array to store folders with custom permissions
        $permissionFolders = @()
        $folderHierarchy = @{}
        $folderDepths = @{}
        
        # Map folder hierarchy
        foreach ($folder in $folderStats) {
            $folderPath = $folder.FolderPath
            $parentPath = Split-Path -Path $folderPath -Parent
            
            if ($parentPath -eq "") {
                $folderDepths[$folderPath] = 1
            }
            else {
                if ($folderHierarchy.ContainsKey($parentPath)) {
                    $folderHierarchy[$parentPath] += $folderPath
                }
                else {
                    $folderHierarchy[$parentPath] = @($folderPath)
                }
                
                $folderDepths[$folderPath] = ($folderDepths[$parentPath] ?? 1) + 1
            }
        }
        
        # Check permissions for each folder
        $maxDepth = 0
        $totalCustomPermissions = 0
        
        foreach ($folder in $folderStats) {
            $folderPath = $folder.FolderPath
            $folderIdentity = "$($mailbox.PrimarySmtpAddress):$($folderPath.Replace('/', '\'))"
            
            try {
                $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction SilentlyContinue
                
                # Filter out default permissions
                $customPermissions = $permissions | Where-Object { 
                    $_.User.DisplayName -ne "Default" -and 
                    $_.User.DisplayName -ne "Anonymous" -and
                    $_.User.DisplayName -ne "Owner" -and
                    $_.AccessRights -ne "None"
                }
                
                if ($customPermissions -and $customPermissions.Count -gt 0) {
                    $depth = $folderDepths[$folderPath] ?? 1
                    
                    if ($depth -gt $maxDepth) {
                        $maxDepth = $depth
                    }
                    
                    $totalCustomPermissions += $customPermissions.Count
                    
                    $permissionFolders += [PSCustomObject]@{
                        FolderPath = $folderPath
                        PermissionCount = $customPermissions.Count
                        Depth = $depth
                        Permissions = $customPermissions | Select-Object @{N="User";E={$_.User.DisplayName}}, AccessRights
                    }
                }
            }
            catch {
                Write-Log -Message "Warning: Could not check permissions for folder $folderPath`: $_" -Level "WARNING"
            }
        }
        
        # Look for complex permission patterns
        $complexPermissionFolders = @()
        $permissionUserCounts = @{}
        
        foreach ($permFolder in $permissionFolders) {
            # Track unique users with permissions
            foreach ($perm in $permFolder.Permissions) {
                $user = $perm.User
                if ($permissionUserCounts.ContainsKey($user)) {
                    $permissionUserCounts[$user]++
                }
                else {
                    $permissionUserCounts[$user] = 1
                }
            }
            
            # Check for deep folders with permissions (depth > 3)
            if ($permFolder.Depth -gt 3) {
                $complexPermissionFolders += $permFolder
            }
            
            # Check for folders with many different permission entries (> 5)
            if ($permFolder.PermissionCount -gt 5) {
                if (-not ($complexPermissionFolders -contains $permFolder)) {
                    $complexPermissionFolders += $permFolder
                }
            }
        }
        
        # Set results
        $Results.HasDeepFolderPermissions = $complexPermissionFolders.Count -gt 0
        $Results.FolderPermissionsCount = $totalCustomPermissions
        $Results.MaxFolderPermissionDepth = $maxDepth
        $Results.ComplexPermissionFolders = $complexPermissionFolders
        
        # Add warnings if complex permissions detected
        if ($Results.HasDeepFolderPermissions) {
            $Results.Warnings += "Mailbox has complex folder permissions that may not migrate correctly"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has complex folder permissions" -Level "WARNING"
            Write-Log -Message "  - $totalCustomPermissions custom permissions across $($permissionFolders.Count) folders" -Level "WARNING"
            Write-Log -Message "  - Maximum folder permission depth: $maxDepth" -Level "WARNING"
            Write-Log -Message "  - $($complexPermissionFolders.Count) folders with complex permission structure" -Level "WARNING"
            
            # List some examples
            foreach ($folder in ($complexPermissionFolders | Select-Object -First 3)) {
                $usersWithAccess = ($folder.Permissions.User -join ", ")
                Write-Log -Message "  - Folder: $($folder.FolderPath) (Depth: $($folder.Depth)), Users: $usersWithAccess" -Level "WARNING"
            }
            
            if ($complexPermissionFolders.Count -gt 3) {
                Write-Log -Message "  - ... and $($complexPermissionFolders.Count - 3) more folders" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Document folder permissions and verify post-migration" -Level "INFO"
        }
        else {
            Write-Log -Message "Mailbox $EmailAddress has $totalCustomPermissions permissions across $($permissionFolders.Count) folders (max depth: $maxDepth)" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to analyze folder permission depth: $_"
        Write-Log -Message "Warning: Failed to analyze folder permission depth for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
