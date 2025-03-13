function Test-NamespaceConflicts {
    <#
    .SYNOPSIS
        Tests for shared namespace conflicts that could affect migration.
    
    .DESCRIPTION
        Checks if a mailbox's domain is part of a shared namespace that could cause
        mail flow or identity matching issues during or after migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-NamespaceConflicts -EmailAddress "user@contoso.com" -Results $results
    
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
    
    if (-not ($script:Config.CheckNamespaceConflicts)) {
        Write-Log -Message "Skipping namespace conflicts check (disabled in config)" -Level "INFO"
        return $true
    }

    try {
        Write-Log -Message "Checking for namespace conflicts: $EmailAddress" -Level "INFO"
        
        # Initialize properties
        $Results | Add-Member -NotePropertyName "HasNamespaceConflict" -NotePropertyValue $false -Force
        $Results | Add-Member -NotePropertyName "NamespaceConflictDetails" -NotePropertyValue $null -Force
        
        # Get mailbox to analyze domains
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        
        # Extract domains from email addresses
        $domains = @()
        foreach ($address in $mailbox.EmailAddresses) {
            if ($address -match "[@](.+)$") {
                $domain = $Matches[1]
                if ($domain -notin $domains) {
                    $domains += $domain
                }
            }
        }
        
        # Get all accepted domains in the organization
        $acceptedDomains = Get-AcceptedDomain
        
        # Check for domain ownership type for each domain
        $conflictingDomains = @()
        
        foreach ($domain in $domains) {
            $acceptedDomain = $acceptedDomains | Where-Object { $_.DomainName -eq $domain }
            
            if ($acceptedDomain) {
                # Check if it's an authoritative domain
                if ($acceptedDomain.DomainType -ne "Authoritative") {
                    $conflictingDomains += [PSCustomObject]@{
                        Domain = $domain
                        DomainType = $acceptedDomain.DomainType
                        IsDefault = $acceptedDomain.Default
                        IsDirSynced = $acceptedDomain.CanHaveCloudCache  # Field indicating if the domain is in a shared namespace
                    }
                }
            }
            elseif ($domain -notlike "*.onmicrosoft.com") {
                # Domain not found in accepted domains but not the default onmicrosoft.com domain
                $conflictingDomains += [PSCustomObject]@{
                    Domain = $domain
                    DomainType = "Unknown"
                    IsDefault = $false
                    IsDirSynced = $false
                }
            }
        }
        
        if ($conflictingDomains.Count -gt 0) {
            $Results.HasNamespaceConflict = $true
            $Results.NamespaceConflictDetails = $conflictingDomains
            
            # Add appropriate warnings
            $sharedDomains = $conflictingDomains | Where-Object { $_.DomainType -eq "InternalRelay" -or $_.IsDirSynced -eq $true }
            $unknownDomains = $conflictingDomains | Where-Object { $_.DomainType -eq "Unknown" }
            
            if ($sharedDomains.Count -gt 0) {
                $domainList = ($sharedDomains.Domain -join ", ")
                $Results.Warnings += "Mailbox has email addresses in shared namespace domains: $domainList"
                
                Write-Log -Message "Warning: Mailbox $EmailAddress has addresses in shared namespace domains: $domainList" -Level "WARNING"
                Write-Log -Message "Recommendation: Verify mail routing for these domains after migration" -Level "INFO"
            }
            
            if ($unknownDomains.Count -gt 0) {
                $domainList = ($unknownDomains.Domain -join ", ")
                $Results.Errors += "Mailbox has email addresses in unverified domains: $domainList"
                $Results.ErrorCodes += "ERR035"
                
                Write-Log -Message "Error: Mailbox $EmailAddress has addresses in unverified domains: $domainList" -Level "ERROR" -ErrorCode "ERR035"
                Write-Log -Message "Recommendation: Verify these domains or remove these email addresses before migration" -Level "INFO"
            }
        }
        else {
            Write-Log -Message "No namespace conflicts found for $EmailAddress" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check for namespace conflicts: $_"
        Write-Log -Message "Warning: Failed to check for namespace conflicts for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
