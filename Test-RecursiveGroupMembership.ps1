function Test-RecursiveGroupMembership {
    <#
    .SYNOPSIS
        Tests for recursive group memberships that might affect migration.
    
    .DESCRIPTION
        Checks if a mailbox user belongs to nested groups that might require
        special attention during migration. Identifies both direct and indirect
        group memberships that may need to be recreated in Exchange Online.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-RecursiveGroupMembership -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking recursive group membership: $EmailAddress" -Level "INFO"
        
        # Get mailbox
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Get direct group memberships
        $directGroups = Get-Recipient -Filter "Members -eq '$($mailbox.DistinguishedName)'" -RecipientTypeDetails 'GroupMailbox','MailUniversalDistributionGroup','MailUniversalSecurityGroup' -ErrorAction SilentlyContinue
        
        $groupInfo = @()
        $nestedGroups = @()
        
        if ($directGroups) {
            foreach ($group in $directGroups) {
                $groupInfo += [PSCustomObject]@{
                    GroupName = $group.DisplayName
                    GroupType = $group.RecipientTypeDetails
                    EmailAddress = $group.PrimarySmtpAddress
                    IsNested = $false
                    NestedLevel = 0
                }
                
                # Use the recursive function to get nested groups
                $nestedGroupMembers = Get-NestedGroupMembership -GroupIdentity $group.DistinguishedName
                
                if ($nestedGroupMembers -and $nestedGroupMembers.Count -gt 0) {
                    foreach ($nestedGroup in $nestedGroupMembers) {
                        $nestedGroups += [PSCustomObject]@{
                            GroupName = $nestedGroup.GroupName
                            GroupType = $nestedGroup.GroupType
                            EmailAddress = $nestedGroup.EmailAddress
                            NestedIn = $nestedGroup.ParentOf
                            IsNested = $true
                            NestedLevel = $nestedGroup.NestedLevel
                        }
                    }
                }
            }
        }
        
        # Combine direct and nested groups
        $allGroups = $groupInfo + $nestedGroups
        
        if ($allGroups.Count -gt 0) {
            $Results.GroupMemberships = $allGroups
            $Results.DirectGroupCount = ($allGroups | Where-Object { -not $_.IsNested }).Count
            $Results.NestedGroupCount = ($allGroups | Where-Object { $_.IsNested }).Count
            
            if ($nestedGroups.Count -gt 0) {
                $Results.Warnings += "User belongs to nested groups which may require special handling during migration"
                $Results.ErrorCodes += "ERR021"
                $maxNestedLevel = ($nestedGroups | Measure-Object -Property NestedLevel -Maximum).Maximum
                
                Write-Log -Message "Warning: User $EmailAddress belongs to nested groups (max depth: $maxNestedLevel):" -Level "WARNING" -ErrorCode "ERR021"
                
                # Log only the first few nested groups to avoid excessive logging
                foreach ($group in ($nestedGroups | Select-Object -First 5)) {
                    Write-Log -Message "  - $($group.GroupName) ($($group.EmailAddress)) nested in $($group.NestedIn) (Level: $($group.NestedLevel))" -Level "WARNING"
                }
                
                if ($nestedGroups.Count -gt 5) {
                    Write-Log -Message "  - ... and $($nestedGroups.Count - 5) more nested groups" -Level "WARNING"
                }
                
                Write-Log -Message "Recommendation: Document group memberships as they may need to be manually recreated" -Level "INFO"
            }
            else {
                Write-Log -Message "User $EmailAddress belongs to $($directGroups.Count) groups (no nested groups)" -Level "INFO"
            }
        }
        else {
            $Results.GroupMemberships = @()
            $Results.DirectGroupCount = 0
            $Results.NestedGroupCount = 0
            Write-Log -Message "No group memberships found for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check group memberships: $_"
        Write-Log -Message "Warning: Failed to check group memberships for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}

# Helper function for recursive group membership resolution
function Get-NestedGroupMembership {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupIdentity,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxDepth = 5,
        
        [Parameter(Mandatory = $false)]
        [int]$CurrentDepth = 0,
        
        [Parameter(Mandatory = $false)]
        [System.Collections.ArrayList]$ProcessedGroups = $null
    )
    
    # Initialize the processed groups tracking array on first call
    if ($null -eq $ProcessedGroups) {
        $ProcessedGroups = New-Object System.Collections.ArrayList
    }
    
    # Prevent infinite recursion by stopping at max depth
    if ($CurrentDepth -ge $MaxDepth) {
        Write-Log -Message "Maximum recursion depth reached for group $GroupIdentity" -Level "WARNING"
        return @()
    }
    
    try {
        # Get the group
        $group = Get-Recipient -Identity $GroupIdentity -ErrorAction Stop
        
        # Add to processed groups to prevent circular references
        if ($ProcessedGroups -notcontains $group.DistinguishedName) {
            [void]$ProcessedGroups.Add($group.DistinguishedName)
        }
        else {
            # Already processed this group, skip to prevent circular reference
            return @()
        }
        
        # Get parent groups that contain this group
        $parentGroups = Get-Recipient -Filter "Members -eq '$($group.DistinguishedName)'" -RecipientTypeDetails 'GroupMailbox','MailUniversalDistributionGroup','MailUniversalSecurityGroup' -ErrorAction SilentlyContinue
        
        $results = @()
        foreach ($parent in $parentGroups) {
            $results += [PSCustomObject]@{
                GroupName = $parent.DisplayName
                GroupType = $parent.RecipientTypeDetails
                EmailAddress = $parent.PrimarySmtpAddress
                NestedLevel = $CurrentDepth + 1
                ParentOf = $group.DisplayName
            }
            
            # Recursively get each parent's parents
            $results += Get-NestedGroupMembership -GroupIdentity $parent.DistinguishedName -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1) -ProcessedGroups $ProcessedGroups
        }
        
        return $results
    }
    catch {
        Write-Log -Message "Error getting nested group membership for $GroupIdentity`: $_" -Level "WARNING"
        return @()
    }
}
