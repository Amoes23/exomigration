function Test-EXOMailboxReadiness {
    <#
    .SYNOPSIS
        Tests if a mailbox is ready for migration between Exchange environments.
    
    .DESCRIPTION
        Performs comprehensive validation of a mailbox to determine if it's ready
        for migration. Supports both migrations to Exchange Online and from
        Exchange Online back to on-premises. Checks various aspects including licensing,
        configurations, permissions, and potential migration blockers.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER ValidationLevel
        Level of validation to perform:
        - Basic: Essential checks only
        - Standard: Basic plus common migration blockers (default)
        - Comprehensive: Full analysis including performance considerations
    
    .PARAMETER MigrationType
        Type of migration to validate for:
        - ToCloud: On-premises to Exchange Online migration (default)
        - FromCloud: Exchange Online to on-premises migration
        - CrossTenant: Exchange Online to Exchange Online (different tenant)
    
    .PARAMETER IncludeInactiveMailboxes
        When specified, includes checking for mailbox inactivity.
    
    .PARAMETER RetryCount
        Number of retry attempts for validation tests.
    
    .PARAMETER RetryDelay
        Delay in seconds between retry attempts.
    
    .PARAMETER SkipExchangeOnlineValidation
        When specified, skips the validation checks against Exchange Online.
        Only use this if you don't have access to Exchange Online during testing.
    
    .PARAMETER SkipOnPremValidation
        When specified, skips the validation checks against on-premises Exchange.
        Only use this if you don't have access to on-premises Exchange during testing.
    
    .EXAMPLE
        Test-EXOMailboxReadiness -EmailAddress "user@contoso.com"
    
    .EXAMPLE
        Test-EXOMailboxReadiness -EmailAddress "user@contoso.com" -ValidationLevel Comprehensive -MigrationType ToCloud
        
    .EXAMPLE
        Test-EXOMailboxReadiness -EmailAddress "user@contoso.com" -MigrationType FromCloud -ValidationLevel Basic
    
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
        [ValidateSet('ToCloud', 'FromCloud', 'CrossTenant')]
        [string]$MigrationType = 'ToCloud',
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeInactiveMailboxes,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 2,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryDelay = 3,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipExchangeOnlineValidation,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipOnPremValidation
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
    
    # Log migration type and validation level
    $migrationDirection = switch ($MigrationType) {
        'ToCloud' { "on-premises to Exchange Online" }
        'FromCloud' { "Exchange Online to on-premises" }
        'CrossTenant' { "Exchange Online to Exchange Online (cross-tenant)" }
    }
    
    Write-Log -Message "Testing migration readiness for: $EmailAddress" -Level "INFO"
    Write-Log -Message "Migration type: $MigrationType ($migrationDirection)" -Level "INFO"
    Write-Log -Message "Validation level: $ValidationLevel" -Level "INFO"

    # Determine which environments to check
    $checkExchangeOnline = -not $SkipExchangeOnlineValidation
    $checkOnPremises = -not $SkipOnPremValidation
    
    # For ToCloud migrations, we need to check both environments
    # For FromCloud migrations, we also need to check both
    # For CrossTenant, we only need to check Exchange Online
    
    if ($MigrationType -eq 'CrossTenant' -and $checkOnPremises) {
        Write-Log -Message "On-premises validation not applicable for cross-tenant migration. Skipping." -Level "INFO"
        $checkOnPremises = $false
    }
    
    # Set OnPremises switch for unified test functions
    $onPremises = $false
    if ($MigrationType -eq 'FromCloud') {
        # When migrating FROM Exchange Online TO on-premises, the target is on-premises
        $onPremises = $true
    }
    
    try {
        # Verify connectivity to required environments
        if ($checkExchangeOnline -and $checkOnPremises) {
            $connectionResult = Connect-MigrationEnvironment -Environment Both
            
            if (-not $connectionResult.ExchangeOnlineConnected -or -not $connectionResult.OnPremisesConnected) {
                if (-not $connectionResult.ExchangeOnlineConnected) {
                    Write-Log -Message "Not connected to Exchange Online. Some validations will be skipped." -Level "WARNING"
                    $results.Warnings += "Exchange Online connection failed. Some validations were skipped."
                    $checkExchangeOnline = $false
                }
                
                if (-not $connectionResult.OnPremisesConnected) {
                    Write-Log -Message "Not connected to on-premises Exchange. Some validations will be skipped." -Level "WARNING"
                    $results.Warnings += "On-premises Exchange connection failed. Some validations were skipped."
                    $checkOnPremises = $false
                }
            }
        }
        elseif ($checkExchangeOnline) {
            $connectionResult = Connect-MigrationEnvironment -Environment ExchangeOnline
            
            if (-not $connectionResult.ExchangeOnlineConnected) {
                Write-Log -Message "Not connected to Exchange Online. Validation cannot proceed." -Level "ERROR"
                $results.Errors += "Not connected to Exchange Online. Validation cannot proceed."
                $results.OverallStatus = "Failed"
                return $results
            }
        }
        elseif ($checkOnPremises) {
            $connectionResult = Connect-MigrationEnvironment -Environment OnPremises
            
            if (-not $connectionResult.OnPremisesConnected) {
                Write-Log -Message "Not connected to on-premises Exchange. Validation cannot proceed." -Level "ERROR"
                $results.Errors += "Not connected to on-premises Exchange. Validation cannot proceed."
                $results.OverallStatus = "Failed"
                return $results
            }
        }
        
        # Define validation test catalog 
        # Define the tests that work the same across environments with the OnPremises parameter
        $commonTests = @(
            @{ Name = 'MailboxFolderStructure'; Function = 'Test-MailboxFolderStructure' },
            @{ Name = 'CalendarAndContactItems'; Function = 'Test-CalendarAndContactItems' },
            @{ Name = 'MailboxStatistics'; Function = 'Test-MailboxStatistics' },
            @{ Name = 'MailboxItemSizeLimits'; Function = 'Test-MailboxItemSizeLimits' },
            @{ Name = 'SpecialMailboxTypes'; Function = 'Test-SpecialMailboxTypes' },
            @{ Name = 'RecoverableItemsSize'; Function = 'Test-RecoverableItemsSize' },
            @{ Name = 'MailboxAuditConfiguration'; Function = 'Test-MailboxAuditConfiguration' },
            @{ Name = 'UnifiedMessagingConfiguration'; Function = 'Test-UnifiedMessagingConfiguration' },
            @{ Name = 'MailboxPermissions'; Function = 'Test-MailboxPermissions' }
        )
        
        # Define tests specific to Exchange Online
        $cloudTests = @(
            @{ Name = 'MailboxLicense'; Function = 'Test-MailboxLicense' },
            @{ Name = 'ADSyncStatus'; Function = 'Test-ADSyncStatus' },
            @{ Name = 'ArchiveMailbox'; Function = 'Test-ArchiveMailbox' },
            @{ Name = 'MoveRequests'; Function = 'Test-MoveRequests' }
        )
        
        # Define tests specific to ToCloud migrations
        $toCloudTests = @(
            @{ Name = 'CloudMailboxPlaceholder'; Function = 'Test-CloudMailboxPlaceholder' },
            @{ Name = 'SoftDeletedMailbox'; Function = 'Test-SoftDeletedMailbox' },
            @{ Name = 'NamespaceConflicts'; Function = 'Test-NamespaceConflicts' },
            @{ Name = 'FolderNameConflicts'; Function = 'Test-FolderNameConflicts' }
        )
        
        # Define tests specific to on-premises (used for FromCloud migrations)
        $onPremTests = @(
            @{ Name = 'OnPremLegacyDN'; Function = 'Test-OnPremLegacyDN' },
            @{ Name = 'JournalConfiguration'; Function = 'Test-JournalConfiguration' }
        )
        
        # Define tests for comprehensive validation only
        $comprehensiveTests = @(
            @{ Name = 'OrphanedPermissions'; Function = 'Test-OrphanedPermissions' },
            @{ Name = 'RecursiveGroupMembership'; Function = 'Test-RecursiveGroupMembership' },
            @{ Name = 'ItemAgeDistribution'; Function = 'Test-ItemAgeDistribution' },
            @{ Name = 'MailboxActivityLevel'; Function = 'Test-MailboxActivityLevel' },
            @{ Name = 'FolderPermissionDepth'; Function = 'Test-FolderPermissionDepth' }
        )
        
        # Determine which tests to run based on migration type and validation level
        $testsToRun = @()
        
        # Common tests are run for all migration types with the appropriate OnPremises parameter
        foreach ($test in $commonTests) {
            if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                $testsToRun += @{
                    Name = $test.Name
                    Function = $test.Function
                    OnPremises = $onPremises
                }
            }
            else {
                $results.Warnings += "Function not found: $($test.Function). This test will be skipped."
                Write-Log -Message "Function not found: $($test.Function). This test will be skipped." -Level "WARNING"
            }
        }
        
        # Add environment-specific tests
        if ($MigrationType -eq 'ToCloud') {
            # Exchange Online tests for target verification
            if ($checkExchangeOnline) {
                foreach ($test in ($cloudTests + $toCloudTests)) {
                    if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                        $testsToRun += @{
                            Name = $test.Name
                            Function = $test.Function
                            OnPremises = $false
                        }
                    }
                }
            }
        }
        elseif ($MigrationType -eq 'FromCloud') {
            # Exchange Online tests for source verification
            if ($checkExchangeOnline) {
                foreach ($test in $cloudTests) {
                    if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                        $testsToRun += @{
                            Name = $test.Name
                            Function = $test.Function
                            OnPremises = $false
                        }
                    }
                }
            }
            
            # On-premises tests for target verification
            if ($checkOnPremises) {
                foreach ($test in $onPremTests) {
                    if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                        $testsToRun += @{
                            Name = $test.Name
                            Function = $test.Function
                            OnPremises = $true
                        }
                    }
                }
            }
        }
        elseif ($MigrationType -eq 'CrossTenant') {
            # Exchange Online tests for both source and target (different tenants)
            if ($checkExchangeOnline) {
                foreach ($test in $cloudTests) {
                    if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                        $testsToRun += @{
                            Name = $test.Name
                            Function = $test.Function
                            OnPremises = $false
                        }
                    }
                }
            }
        }
        
        # Add comprehensive tests if the validation level is set to Comprehensive
        if ($ValidationLevel -eq 'Comprehensive') {
            foreach ($test in $comprehensiveTests) {
                if (Get-Command -Name $test.Function -ErrorAction SilentlyContinue) {
                    $testsToRun += @{
                        Name = $test.Name
                        Function = $test.Function
                        OnPremises = $onPremises
                    }
                }
            }
        }
        
        # Run all selected tests
        $testResults = @{}
        
        foreach ($test in $testsToRun) {
            try {
                # Create parameter set for the function
                $params = @{
                    EmailAddress = $EmailAddress
                    Results = $results
                }
                
                # Add OnPremises switch if specified
                if ($test.OnPremises) {
                    $params.Add('OnPremises', $true)
                }
                
                # Use retry logic for each test
                $testResults[$test.Name] = Invoke-WithRetry -ScriptBlock { 
                    & $test.Function @params
                } -MaxRetries $RetryCount -DelaySeconds $RetryDelay -Activity "$($test.Name) validation"
                
                if ($testResults[$test.Name]) {
                    Write-Log -Message "$($test.Name) validation passed for $EmailAddress" -Level "DEBUG"
                }
                else {
                    Write-Log -Message "$($test.Name) validation failed for $EmailAddress" -Level "WARNING"
                }
            }
            catch {
                $testResults[$test.Name] = $false
                $errorLevel = if ($ValidationLevel -eq 'Comprehensive' -and $test.Name -in $comprehensiveTests.Name) {
                    "WARNING" # Only add warnings for comprehensive tests failing
                } else {
                    "ERROR"   # For basic and standard tests, errors are more severe
                }
                
                Write-Log -Message "$($test.Name) validation error for $EmailAddress`: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level $errorLevel
                
                if ($errorLevel -eq "ERROR") {
                    $results.Errors += "Failed to perform $($test.Name) validation: $($_.Exception.Message)"
                }
                else {
                    $results.Warnings += "Failed to perform $($test.Name) validation: $($_.Exception.Message)"
                }
            }
        }
        
        # Check for inactive mailboxes if requested
        if ($IncludeInactiveMailboxes) {
            try {
                if (Get-Command -Name "Test-MailboxActivityLevel" -ErrorAction SilentlyContinue) {
                    # Use the unified function
                    Test-MailboxActivityLevel -EmailAddress $EmailAddress -Results $results -OnPremises:$onPremises
                }
                else {
                    Write-Log -Message "Function Test-MailboxActivityLevel not found, skipping inactivity check" -Level "WARNING"
                    $results.Warnings += "Inactive mailbox check skipped: Function not found"
                }
            }
            catch {
                Write-Log -Message "Inactive mailbox check failed for $EmailAddress`: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
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
            Write-Log -Message "Migration readiness test completed with WARNINGS for $EmailAddress`: $($results.Warnings.Count) warnings found" -Level "WARNING"
        }
        else {
            $results.OverallStatus = "Ready"
            Write-Log -Message "Migration readiness test PASSED for $EmailAddress - Mailbox is READY for $migrationDirection migration" -Level "SUCCESS"
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
