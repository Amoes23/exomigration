function Test-EXOMailboxReadiness {
    <#
    .SYNOPSIS
        Tests if a mailbox is ready for migration to Exchange Online.
    
    .DESCRIPTION
        Performs comprehensive validation of a mailbox to determine if it's ready
        for migration to Exchange Online. Checks various aspects including licensing,
        configurations, permissions, and potential migration blockers.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER ValidationLevel
        Level of validation to perform:
        - Basic: Essential checks only
        - Standard: Basic plus common migration blockers (default)
        - Comprehensive: Full analysis including performance considerations
    
    .PARAMETER IncludeInactiveMailboxes
        When specified, includes checking for mailbox inactivity.
    
    .PARAMETER RetryCount
        Number of retry attempts for validation tests.
    
    .PARAMETER RetryDelay
        Delay in seconds between retry attempts.
        
    .PARAMETER OnPremisesMigration
        When specified, treats this as an on-premises to Exchange Online migration,
        validating that the mailbox exists on-premises and not in Exchange Online.
    
    .PARAMETER SkipExchangeOnlineValidation
        When specified, skips the validation checks against Exchange Online.
        Only use this if you don't have access to Exchange Online during testing.
    
    .EXAMPLE
        Test-EXOMailboxReadiness -EmailAddress "user@contoso.com"
    
    .EXAMPLE
        Test-EXOMailboxReadiness -EmailAddress "user@contoso.com" -ValidationLevel Comprehensive
        
    .EXAMPLE
        Test-EXOMailboxReadiness -EmailAddress "user@contoso.com" -OnPremisesMigration -ValidationLevel Basic
    
    .OUTPUTS
        PSCustomObject with validation results and readiness status.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Basic', 'Standard', 'Comprehensive')]
        [string]$ValidationLevel = 'Standard',
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeInactiveMailboxes,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 2,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryDelay = 3,
        
        [Parameter(Mandatory = $false)]
        [switch]$OnPremisesMigration,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipExchangeOnlineValidation
    )
    
    # Create a result object with default properties
    try {
        $results = New-MailboxTestResult -EmailAddress $EmailAddress
        $results.ValidationLevel = $ValidationLevel
    }
    catch {
        Write-Log -Message "Failed to create result object: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
        throw "Unable to initialize validation. Error: $($_.Exception.Message)"
    }
    
    Write-Log -Message "Testing migration readiness for: $EmailAddress with validation level: $ValidationLevel" -Level "INFO"
    if ($OnPremisesMigration) {
        Write-Log -Message "Mode: On-premises to Exchange Online migration" -Level "INFO"
    }
    
    try {
        # For on-premises migrations, first check the on-premises mailbox
        if ($OnPremisesMigration) {
            try {
                # First check if we're connected to on-premises Exchange
                try {
                    $null = Get-ExchangeServer -ErrorAction Stop
                    Write-Log -Message "Connected to on-premises Exchange" -Level "INFO"
                }
                catch {
                    Write-Log -Message "Not connected to on-premises Exchange. Please connect to on-premises Exchange first." -Level "WARNING"
                    $results.Warnings += "Not connected to on-premises Exchange. Source mailbox validation may be limited."
                }
                
                # Try to get on-premises mailbox
                try {
                    $onpremMailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
                    Write-Log -Message "On-premises mailbox found: $($onpremMailbox.DisplayName)" -Level "INFO"
                    
                    # Copy basic info to results
                    $results.DisplayName = $onpremMailbox.DisplayName
                    $results.RecipientType = $onpremMailbox.RecipientType
                    $results.RecipientTypeDetails = $onpremMailbox.RecipientTypeDetails
                    $results.MailboxEnabled = $true
                    $results.ExistingMailbox = $false  # Not in Exchange Online yet, which is expected
                }
                catch {
                    Write-Log -Message "On-premises mailbox not found: $EmailAddress" -Level "ERROR"
                    $results.Errors += "On-premises source mailbox not found. Please verify the email address is correct."
                    $results.OverallStatus = "Failed"
                    return $results
                }
            }
            catch {
                Write-Log -Message "Error during on-premises mailbox validation: $($_.Exception.Message)" -Level "ERROR"
                $results.Errors += "Failed to validate on-premises mailbox: $($_.Exception.Message)"
                $results.OverallStatus = "Failed" 
                return $results
            }
        }
        
        # Verify we're connected to Exchange Online and Microsoft Graph before proceeding
        # Skip if OnPremisesMigration & SkipExchangeOnlineValidation are both set
        if (-not ($OnPremisesMigration -and $SkipExchangeOnlineValidation)) {
            try {
                $null = Get-AcceptedDomain -ErrorAction Stop
            }
            catch {
                Write-Log -Message "Not connected to Exchange Online. Please run Connect-ExchangeOnline first." -Level "ERROR"
                $results.Errors += "Not connected to Exchange Online. Please run Connect-ExchangeOnline first."
                $results.OverallStatus = "Failed"
                return $results
            }
            
            try {
                $null = Get-MgUser -Top 1 -ErrorAction Stop
            }
            catch {
                Write-Log -Message "Not connected to Microsoft Graph. Please run Connect-MgGraph first." -Level "WARNING"
                $results.Warnings += "Not connected to Microsoft Graph. Some validations may be skipped."
            }
            
            # For on-premises migrations, check if the mailbox already exists in Exchange Online
            # which would be a conflict and potential issue
            if ($OnPremisesMigration) {
                try {
                    $existingEXOMailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
                    if ($existingEXOMailbox) {
                        Write-Log -Message "Mailbox already exists in Exchange Online: $EmailAddress" -Level "ERROR"
                        $results.Errors += "Mailbox already exists in Exchange Online. Migration not possible without removing the existing mailbox."
                        $results.OverallStatus = "Failed"
                        return $results
                    }
                    else {
                        Write-Log -Message "No existing mailbox found in Exchange Online for $EmailAddress (expected for migration)" -Level "INFO"
                    }
                }
                catch {
                    # This is expected - mailbox shouldn't exist in Exchange Online yet
                    Write-Log -Message "No existing mailbox found in Exchange Online for $EmailAddress (expected for migration)" -Level "INFO"
                }
            }
        }
        
        # Determine which validation tests to run based on migration type
        if ($OnPremisesMigration) {
            # For on-premises migration, focus on readiness and conflict checks
            if (-not $SkipExchangeOnlineValidation) {
                # Only run these if we're checking against Exchange Online
                $basicTests = @(
                    @{ Name = 'AliasConflicts'; Function = 'Test-AliasConflicts' },
                    @{ Name = 'SoftDeletedMailbox'; Function = 'Test-SoftDeletedMailbox' },
                    @{ Name = 'CloudMailboxPlaceholder'; Function = 'Test-CloudMailboxPlaceholder' }
                )
            } else {
                $basicTests = @()
            }
            
            # If we have on-premises connection, add these tests
            if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) {
                # Basic on-premises tests that should always be run
                $onPremBasicTests = @(
                    @{ Name = 'OnPremMailboxConfiguration'; Function = 'Test-OnPremMailboxConfiguration' },
                    @{ Name = 'OnPremPermissions'; Function = 'Test-OnPremPermissions' },
                    @{ Name = 'OnPremLegacyDN'; Function = 'Test-OnPremLegacyDN' }
                )
                $basicTests += $onPremBasicTests
            }
        } else {
            # Original tests for cloud mailbox validation
            $basicTests = @(
                @{ Name = 'MailboxConfiguration'; Function = 'Test-MailboxConfiguration' },
                @{ Name = 'MailboxLicense'; Function = 'Test-MailboxLicense' },
                @{ Name = 'MoveRequests'; Function = 'Test-MoveRequests' }
            )
            
            # Add the more advanced tests if they're available
            $advancedBasicTests = @(
                @{ Name = 'ADSyncStatus'; Function = 'Test-ADSyncStatus' },
                @{ Name = 'CloudMailboxPlaceholder'; Function = 'Test-CloudMailboxPlaceholder' },
                @{ Name = 'SoftDeletedMailbox'; Function = 'Test-SoftDeletedMailbox' },
                @{ Name = 'AliasConflicts'; Function = 'Test-AliasConflicts' }
            )
            
            # Only add tests that actually exist as commands
            foreach ($test in $advancedBasicTests) {
                if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                    $basicTests += $test
                }
            }
        }
        
        # Execute the tests
        foreach ($test in $basicTests) {
            try {
                # Check if function exists before invoking
                if (-not (Get-Command -Name $test.Function -ErrorAction SilentlyContinue)) {
                    Write-Log -Message "Validation function not found: $($test.Function). This test will be skipped." -Level "WARNING"
                    $results.Warnings += "Validation function not found: $($test.Function)"
                    continue
                }
                
                # Use retry logic for each test
                Invoke-WithRetry -ScriptBlock { 
                    & $test.Function -EmailAddress $EmailAddress -Results $results 
                } -MaxRetries $RetryCount -DelaySeconds $RetryDelay -Activity "$($test.Name) validation"
                
                Write-Log -Message "$($test.Name) validation passed for $EmailAddress" -Level "DEBUG"
            }
            catch {
                Write-Log -Message "$($test.Name) validation failed for $EmailAddress $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
                if ($_.ScriptStackTrace) {
                    Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
                }
                $results.Errors += "Failed to perform $($test.Name) validation: $($_.Exception.Message)"
            }
        }
        
        # Skip standard and comprehensive tests for on-premises migrations unless explicitly implemented
        # This prevents errors from running cloud-focused tests on on-premises mailboxes
        if ($OnPremisesMigration) {
            # You can implement on-premises specific standard and comprehensive tests here
            # For now, skip to maintain compatibility
            Write-Log -Message "Skipping standard and comprehensive tests for on-premises migration" -Level "INFO"
            
            # You could implement on-prem specific tests like:
            if ($ValidationLevel -in @('Standard', 'Comprehensive')) {
                $onPremStandardTests = @(
                    @{ Name = 'OnPremMailboxStatistics'; Function = 'Test-OnPremMailboxStatistics' },
                    @{ Name = 'OnPremItemSizeLimits'; Function = 'Test-OnPremItemSizeLimits' }
                )
                
                foreach ($test in $onPremStandardTests) {
                    if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                        try {
                            Invoke-WithRetry -ScriptBlock { 
                                & $test.Function -EmailAddress $EmailAddress -Results $results 
                            } -MaxRetries $RetryCount -DelaySeconds $RetryDelay -Activity "$($test.Name) validation"
                            
                            Write-Log -Message "$($test.Name) validation passed for $EmailAddress" -Level "DEBUG"
                        }
                        catch {
                            Write-Log -Message "$($test.Name) validation failed for $EmailAddress: $($_.Exception.Message)" -Level "WARNING"
                            $results.Warnings += "Failed to perform $($test.Name) validation: $($_.Exception.Message)"
                        }
                    }
                }
            }
        }
        else {
            # Standard validation - Performed for 'Standard' and 'Comprehensive' levels
            if ($ValidationLevel -in @('Standard', 'Comprehensive')) {
                $standardTests = @(
                    @{ Name = 'MailboxStatistics'; Function = 'Test-MailboxStatistics' },
                    @{ Name = 'MailboxPermissions'; Function = 'Test-MailboxPermissions' },
                    @{ Name = 'MailboxItemSizeLimits'; Function = 'Test-MailboxItemSizeLimits' },
                    @{ Name = 'SpecialMailboxTypes'; Function = 'Test-SpecialMailboxTypes' }
                )
                
                # Add advanced tests if they exist
                $advancedStandardTests = @(
                    @{ Name = 'ArchiveMailbox'; Function = 'Test-ArchiveMailbox' },
                    @{ Name = 'RecoverableItemsSize'; Function = 'Test-RecoverableItemsSize' },
                    @{ Name = 'NamespaceConflicts'; Function = 'Test-NamespaceConflicts' },
                    @{ Name = 'JournalConfiguration'; Function = 'Test-JournalConfiguration' }
                )
                
                foreach ($test in $advancedStandardTests) {
                    if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                        $standardTests += $test
                    }
                }
                
                foreach ($test in $standardTests) {
                    try {
                        # Check if function exists before invoking
                        if (-not (Get-Command -Name $test.Function -ErrorAction SilentlyContinue)) {
                            Write-Log -Message "Validation function not found: $($test.Function)" -Level "WARNING"
                            $results.Warnings += "Validation function not found: $($test.Function)"
                            continue
                        }
                        
                        Invoke-WithRetry -ScriptBlock { 
                            & $test.Function -EmailAddress $EmailAddress -Results $results 
                        } -MaxRetries $RetryCount -DelaySeconds $RetryDelay -Activity "$($test.Name) validation"
                        
                        Write-Log -Message "$($test.Name) validation passed for $EmailAddress" -Level "DEBUG"
                    }
                    catch {
                        Write-Log -Message "$($test.Name) validation failed for $EmailAddress $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
                        if ($_.ScriptStackTrace) {
                            Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
                        }
                        $results.Warnings += "Failed to perform $($test.Name) validation: $($_.Exception.Message)"
                    }
                }
            }
            
            # Comprehensive validation - Only performed for 'Comprehensive' level
            if ($ValidationLevel -eq 'Comprehensive') {
                $comprehensiveTests = @(
                    @{ Name = 'UnifiedMessagingConfiguration'; Function = 'Test-UnifiedMessagingConfiguration' },
                    @{ Name = 'OrphanedPermissions'; Function = 'Test-OrphanedPermissions' },
                    @{ Name = 'RecursiveGroupMembership'; Function = 'Test-RecursiveGroupMembership' },
                    @{ Name = 'MailboxFolderStructure'; Function = 'Test-MailboxFolderStructure' },
                    @{ Name = 'CalendarAndContactItems'; Function = 'Test-CalendarAndContactItems' }
                )
                
                # Add advanced tests if they exist
                $advancedComprehensiveTests = @(
                    @{ Name = 'MailboxAuditConfiguration'; Function = 'Test-MailboxAuditConfiguration' },
                    @{ Name = 'FolderNameConflicts'; Function = 'Test-FolderNameConflicts' },
                    @{ Name = 'ItemAgeDistribution'; Function = 'Test-ItemAgeDistribution' },
                    @{ Name = 'MailboxActivityLevel'; Function = 'Test-MailboxActivityLevel' },
                    @{ Name = 'FolderPermissionDepth'; Function = 'Test-FolderPermissionDepth' }
                )
                
                foreach ($test in $advancedComprehensiveTests) {
                    if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                        $comprehensiveTests += $test
                    }
                }
                
                foreach ($test in $comprehensiveTests) {
                    try {
                        # Check if function exists before invoking
                        if (-not (Get-Command -Name $test.Function -ErrorAction SilentlyContinue)) {
                            Write-Log -Message "Validation function not found: $($test.Function)" -Level "WARNING"
                            $results.Warnings += "Validation function not found: $($test.Function)"
                            continue
                        }
                        
                        Invoke-WithRetry -ScriptBlock { 
                            & $test.Function -EmailAddress $EmailAddress -Results $results 
                        } -MaxRetries $RetryCount -DelaySeconds $RetryDelay -Activity "$($test.Name) validation"
                        
                        Write-Log -Message "$($test.Name) validation passed for $EmailAddress" -Level "DEBUG"
                    }
                    catch {
                        # Comprehensive tests failing shouldn't block migration, just add warnings
                        Write-Log -Message "$($test.Name) validation failed for $EmailAddress $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
                        if ($_.ScriptStackTrace) {
                            Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
                        }
                        $results.Warnings += "Failed to perform $($test.Name) validation: $($_.Exception.Message)"
                    }
                }
            }
        }
        
        # Check for inactive mailboxes if requested
        if ($IncludeInactiveMailboxes) {
            try {
                if (Get-Command -Name "Test-InactiveMailboxes" -ErrorAction SilentlyContinue) {
                    Test-InactiveMailboxes -EmailAddress $EmailAddress -Results $results
                }
                else {
                    Write-Log -Message "Function Test-InactiveMailboxes not found, skipping inactivity check" -Level "WARNING"
                    $results.Warnings += "Inactive mailbox check skipped: Function not found"
                }
            }
            catch {
                Write-Log -Message "Inactive mailbox check failed for $EmailAddress $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
                $results.Warnings += "Failed to check mailbox activity: $($_.Exception.Message)"
            }
        }
        
        # Determine overall status
        if ($results.Errors.Count -gt 0) {
            $results.OverallStatus = "Failed"
            Write-Log -Message "Migration readiness test FAILED for $EmailAddress with $($results.Errors.Count) errors" -Level "ERROR"
        }
        elseif ($results.Warnings.Count -gt 0) {
            $results.OverallStatus = "Warning"
            Write-Log -Message "Migration readiness test completed with WARNINGS for $EmailAddress : $($results.Warnings.Count) warnings found" -Level "WARNING"
        }
        else {
            $results.OverallStatus = "Ready"
            Write-Log -Message "Migration readiness test PASSED for $EmailAddress - Mailbox is READY for migration" -Level "SUCCESS"
        }
        
        return $results
    }
    catch {
        $results.Errors += "Test failed: $($_.Exception.Message)"
        $results.OverallStatus = "Failed"
        Write-Log -Message "Migration readiness test failed for $EmailAddress`: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
        if ($_.ScriptStackTrace) {
            Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
        }
        return $results
    }
}