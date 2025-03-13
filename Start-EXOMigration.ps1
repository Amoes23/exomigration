function Start-EXOMigration {
    <#
    .SYNOPSIS
        Initiates a mailbox migration from on-premises Exchange to Exchange Online.
    
    .DESCRIPTION
        Automates the process of migrating mailboxes from on-premises Exchange to Exchange Online.
        Performs pre-migration validation, creates migration batches, and generates detailed reports.
        Includes checkpointing for resuming interrupted migrations.
    
    .PARAMETER BatchFilePath
        Path to the CSV file containing mailboxes to migrate. Must have "EmailAddress" as a header.
    
    .PARAMETER ConfigPath
        Path to the JSON configuration file. If not specified, defaults to .\Config\ExchangeMigrationConfig.json
    
    .PARAMETER ValidationLevel
        Level of validation to perform:
        - Basic: Essential checks only (licensing, connectivity)
        - Standard: Basic plus common migration blockers (default)
        - Comprehensive: Full analysis including performance considerations
    
    .PARAMETER MaxConcurrentMailboxes
        Maximum number of mailboxes to process in parallel. Default is 5.
    
    .PARAMETER BatchSize
        Number of mailboxes per processing batch for memory optimization. Default is 100.
    
    .PARAMETER DryRun
        If specified, validates everything but doesn't create the migration batch.
    
    .PARAMETER Force
        If specified, attempts to create the migration batch even if validation errors are present.
    
    .PARAMETER Resume
        Resumes a previously interrupted migration process.
    
    .PARAMETER ResumeFromState
        Path to a specific state file to resume from.
    
    .PARAMETER SkipValidation
        Skips the validation phase and uses existing results.
    
    .PARAMETER IncludeWarnings
        Includes mailboxes with warnings in the migration batch.
    
    .EXAMPLE
        Start-EXOMigration -BatchFilePath ".\Batches\Finance.csv"
    
    .EXAMPLE
        Start-EXOMigration -BatchFilePath ".\Batches\Sales.csv" -ValidationLevel Comprehensive -DryRun
    
    .EXAMPLE
        Start-EXOMigration -BatchFilePath ".\Batches\HR.csv" -Resume
    
    .OUTPUTS
        [PSCustomObject] Returns a custom object with migration results and status.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$BatchFilePath,
        
        [Parameter(Mandatory = $false)]
        [ValidateScript({Test-Path $_ -PathType Leaf -ErrorAction SilentlyContinue -or $_ -eq ".\Config\ExchangeMigrationConfig.json"})]
        [string]$ConfigPath = ".\Config\ExchangeMigrationConfig.json",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Basic', 'Standard', 'Comprehensive')]
        [string]$ValidationLevel = 'Standard',
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 20)]
        [int]$MaxConcurrentMailboxes = 5,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 500)]
        [int]$BatchSize = 100,
        
        [Parameter(Mandatory = $false)]
        [switch]$DryRun,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force,
        
        [Parameter(Mandatory = $false)]
        [switch]$Resume,
        
        [Parameter(Mandatory = $false)]
        [ValidateScript({Test-Path $_ -PathType Leaf -ErrorAction SilentlyContinue})]
        [string]$ResumeFromState,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipValidation,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeWarnings
    )
    
    begin {
        # Script start information
        $scriptStartTime = Get-Date
        $scriptName = $MyInvocation.MyCommand.Name
        
        Write-Host "====================================================" -ForegroundColor Cyan
        Write-Host "  Exchange Online Migration Tool" -ForegroundColor Cyan
        Write-Host "  Version: $script:ScriptVersion" -ForegroundColor Cyan
        Write-Host "  Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")" -ForegroundColor Cyan
        Write-Host "====================================================" -ForegroundColor Cyan
        Write-Host ""
        
        # Generate a unique migration ID for this run
        $migrationId = [guid]::NewGuid().ToString()
        
        # Initialize result object
        $result = [PSCustomObject]@{
            Success = $false
            DryRun = $DryRun
            BatchCreated = $false
            BatchId = $null
            BatchStatus = $null
            MigratedMailboxCount = 0
            ReadyMailboxes = 0
            WarningMailboxes = 0
            FailedMailboxes = 0
            IncludedWarnings = $false
            ReportPath = $null
            ElapsedTime = $null
            Error = $null
            ErrorDetails = $null
        }
    }
    
    process {
        try {
            # Load configuration
            Write-Host "Loading configuration..." -NoNewline
            $script:Config = Import-MigrationConfig -ConfigPath $ConfigPath
            if (-not $script:Config) {
                throw "Failed to load configuration from $ConfigPath"
            }
            Write-Host "Done" -ForegroundColor Green
            
            # Override config settings with parameters if provided
            if ($PSBoundParameters.ContainsKey('ValidationLevel')) {
                $script:Config.ValidationLevel = $ValidationLevel
            }
            
            if ($PSBoundParameters.ContainsKey('MaxConcurrentMailboxes')) {
                $script:Config.MaxConcurrentMailboxes = $MaxConcurrentMailboxes
            }
            
            if ($PSBoundParameters.ContainsKey('BatchSize')) {
                $script:Config.BatchSize = $BatchSize
            }
            
            # Prepare state file path for checkpointing
            $batchName = [System.IO.Path]::GetFileNameWithoutExtension($BatchFilePath)
            $stateFilePath = Join-Path -Path $script:Config.LogPath -ChildPath "$batchName-state.json"
            
            # Prepare workspace paths
            $workspacePath = Join-Path -Path $script:Config.WorkspacePath -ChildPath "Workspace-$batchName"
            $validationResultsPath = Join-Path -Path $workspacePath -ChildPath "ValidationResults"
            $reportPath = Join-Path -Path $script:Config.ReportPath -ChildPath "$batchName-reports"
            
            # Create workspace directories
            foreach ($path in @($workspacePath, $validationResultsPath, $reportPath)) {
                if (-not (Test-Path -Path $path)) {
                    New-Item -ItemType Directory -Path $path -Force | Out-Null
                    Write-Verbose "Created directory: $path"
                }
            }
            
            # Initialize logging
            $logPath = Join-Path -Path $script:Config.LogPath -ChildPath "$batchName-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
            $script:LogFile = $logPath
            
            Write-Log -Message "Starting Exchange Online Migration process for batch: $batchName" -Level "INFO"
            Write-Log -Message "Migration ID: $migrationId" -Level "INFO"
            Write-Log -Message "Using configuration from: $ConfigPath" -Level "INFO"
            
            # Verify disk space
            $diskSpaceCheck = Test-DiskSpace -Paths @(
                $script:Config.LogPath, 
                $script:Config.ReportPath, 
                $workspacePath
            )
            
            if (-not $diskSpaceCheck) {
                throw "Insufficient disk space for migration operation. See log for details."
            }
            
            # Migration state object to track progress
            $migrationState = [PSCustomObject]@{
                MigrationId = $migrationId
                BatchName = $batchName
                BatchFilePath = $BatchFilePath
                StartTime = $scriptStartTime
                LastUpdated = Get-Date
                CurrentStage = "Initializing"
                ValidationCompleted = $false
                ValidationResultsPath = $null
                ReadyMailboxes = @()
                WarningMailboxes = @()
                FailedMailboxes = @()
                BatchCreated = $false
                BatchId = $null
                CompletionTime = $null
            }
            
            # Handle resume scenario
            if ($Resume -or $ResumeFromState) {
                $stateFileToLoad = if ($ResumeFromState) { $ResumeFromState } else { $stateFilePath }
                
                if (Test-Path -Path $stateFileToLoad) {
                    Write-Log -Message "Resuming migration from state file: $stateFileToLoad" -Level "INFO"
                    
                    try {
                        $previousState = Get-Content -Path $stateFileToLoad -Raw | ConvertFrom-Json
                        
                        # Validate the state file
                        if (-not $previousState.MigrationId -or -not $previousState.BatchName) {
                            throw "Invalid state file format"
                        }
                        
                        # Update our current state with the previous one
                        $migrationState = $previousState
                        $migrationState.LastUpdated = Get-Date
                        
                        Write-Log -Message "Resumed migration state from: $($migrationState.CurrentStage)" -Level "INFO"
                        Write-Log -Message "Original start time: $($migrationState.StartTime)" -Level "INFO"
                        
                        # If validation was completed, we can skip it
                        if ($migrationState.ValidationCompleted -and $migrationState.ValidationResultsPath -and (Test-Path -Path $migrationState.ValidationResultsPath)) {
                            $SkipValidation = $true
                            Write-Log -Message "Previous validation results found, skipping validation" -Level "INFO"
                        }
                    }
                    catch {
                        Write-Log -Message "Failed to resume from state file: $_" -Level "ERROR"
                        Write-Log -Message "Starting fresh migration process" -Level "WARNING"
                        
                        # Backup the corrupted state file
                        $backupPath = "$stateFileToLoad.backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
                        Copy-Item -Path $stateFileToLoad -Destination $backupPath -Force
                        Write-Log -Message "Backed up corrupted state file to: $backupPath" -Level "INFO"
                    }
                }
                else {
                    Write-Log -Message "State file not found: $stateFileToLoad" -Level "WARNING"
                    Write-Log -Message "Starting fresh migration process" -Level "WARNING"
                }
            }
            
            # 1. Check for dependencies
            if ($migrationState.CurrentStage -eq "Initializing") {
                Write-Log -Message "Checking dependencies..." -Level "INFO"
                $migrationState.CurrentStage = "CheckingDependencies"
                Save-MigrationState -State $migrationState -Path $stateFilePath
                
                $dependenciesOK = Test-MigrationDependencies
                if (-not $dependenciesOK) {
                    throw "Required dependencies are missing"
                }
                
                Write-Log -Message "All dependencies verified" -Level "SUCCESS"
                $migrationState.CurrentStage = "ConnectingServices"
                Save-MigrationState -State $migrationState -Path $stateFilePath
            }
            
            # 2. Connect to required services
            if ($migrationState.CurrentStage -eq "ConnectingServices") {
                Write-Log -Message "Connecting to migration services..." -Level "INFO"
                
                $connected = Connect-MigrationServices
                if (-not $connected) {
                    throw "Failed to connect to migration services"
                }
                
                Write-Log -Message "Successfully connected to all required services" -Level "SUCCESS"
                $migrationState.CurrentStage = "ValidatingMailboxes"
                Save-MigrationState -State $migrationState -Path $stateFilePath
            }
            
            # 3. Validate mailboxes
            if ($migrationState.CurrentStage -eq "ValidatingMailboxes" -and -not $SkipValidation) {
                Write-Log -Message "Starting mailbox validation with level: $ValidationLevel" -Level "INFO"
                
                # Validate mailboxes with batching for memory efficiency
                $validationResultsFile = Invoke-EXOParallelValidation -BatchFilePath $BatchFilePath `
                    -ThrottleLimit $script:Config.MaxConcurrentMailboxes `
                    -BatchSize $script:Config.BatchSize `
                    -ValidationLevel $script:Config.ValidationLevel `
                    -OutputPath $validationResultsPath
                
                if (-not $validationResultsFile -or -not (Test-Path -Path $validationResultsFile)) {
                    throw "Mailbox validation failed or returned no results"
                }
                
                # Load validation results
                $validationResults = Import-Clixml -Path $validationResultsFile
                
                # Update the migration state
                $migrationState.ValidationCompleted = $true
                $migrationState.ValidationResultsPath = $validationResultsFile
                $migrationState.ReadyMailboxes = @($validationResults | Where-Object { $_.OverallStatus -eq "Ready" } | Select-Object -ExpandProperty EmailAddress)
                $migrationState.WarningMailboxes = @($validationResults | Where-Object { $_.OverallStatus -eq "Warning" } | Select-Object -ExpandProperty EmailAddress)
                $migrationState.FailedMailboxes = @($validationResults | Where-Object { $_.OverallStatus -eq "Failed" } | Select-Object -ExpandProperty EmailAddress)
                
                Write-Log -Message "Mailbox validation complete:" -Level "SUCCESS"
                Write-Log -Message "  - Ready: $($migrationState.ReadyMailboxes.Count)" -Level "INFO"
                Write-Log -Message "  - Warning: $($migrationState.WarningMailboxes.Count)" -Level "INFO"
                Write-Log -Message "  - Failed: $($migrationState.FailedMailboxes.Count)" -Level "INFO"
                
                $migrationState.CurrentStage = "GeneratingReport"
                Save-MigrationState -State $migrationState -Path $stateFilePath
            }
            elseif ($SkipValidation -and $migrationState.ValidationCompleted) {
                Write-Log -Message "Skipping validation as requested or resumed from previous run" -Level "INFO"
                Write-Log -Message "Using validation results from: $($migrationState.ValidationResultsPath)" -Level "INFO"
                
                if (-not (Test-Path -Path $migrationState.ValidationResultsPath)) {
                    throw "Validation results file not found: $($migrationState.ValidationResultsPath)"
                }
                
                $validationResults = Import-Clixml -Path $migrationState.ValidationResultsPath
                
                Write-Log -Message "Loaded previous validation results:" -Level "INFO"
                Write-Log -Message "  - Ready: $($migrationState.ReadyMailboxes.Count)" -Level "INFO"
                Write-Log -Message "  - Warning: $($migrationState.WarningMailboxes.Count)" -Level "INFO"
                Write-Log -Message "  - Failed: $($migrationState.FailedMailboxes.Count)" -Level "INFO"
                
                $migrationState.CurrentStage = "GeneratingReport"
                Save-MigrationState -State $migrationState -Path $stateFilePath
            }
            else {
                throw "Invalid state: Validation not completed and SkipValidation not specified"
            }
            
            # 4. Generate HTML report
            if ($migrationState.CurrentStage -eq "GeneratingReport") {
                Write-Log -Message "Generating HTML report..." -Level "INFO"
                
                $reportFilePath = New-EXOMigrationReport -ValidationResults $validationResults -ReportPath $reportPath -BatchName $batchName
                
                if (-not $reportFilePath -or -not (Test-Path -Path $reportFilePath)) {
                    Write-Log -Message "Failed to generate HTML report" -Level "ERROR"
                }
                else {
                    Write-Log -Message "HTML report generated: $reportFilePath" -Level "SUCCESS"
                    $result.ReportPath = $reportFilePath
                    
                    # Try to open the report
                    try {
                        Start-Process -FilePath $reportFilePath
                    }
                    catch {
                        Write-Log -Message "Could not open report: $_" -Level "WARNING"
                    }
                }
                
                $migrationState.CurrentStage = "PrepareForBatchCreation"
                Save-MigrationState -State $migrationState -Path $stateFilePath
            }
            
            # 5. Prepare for batch creation or exit if dry run
            if ($migrationState.CurrentStage -eq "PrepareForBatchCreation") {
                if ($DryRun) {
                    Write-Log -Message "Dry run mode: No migration batch will be created" -Level "WARNING"
                    Write-Log -Message "Migration preparation completed successfully in dry run mode" -Level "SUCCESS"
                    
                    $migrationState.CurrentStage = "Completed"
                    $migrationState.CompletionTime = Get-Date
                    Save-MigrationState -State $migrationState -Path $stateFilePath
                    
                    # Update result object
                    $result.Success = $true
                    $result.ReadyMailboxes = $migrationState.ReadyMailboxes.Count
                    $result.WarningMailboxes = $migrationState.WarningMailboxes.Count
                    $result.FailedMailboxes = $migrationState.FailedMailboxes.Count
                    $result.ElapsedTime = (Get-Date) - $scriptStartTime
                    
                    # Return early with dry run results
                    return $result
                }
                
                # If not in dry run mode, prompt for confirmation unless Force is specified
                $includeMailboxesWithWarnings = $false
                
                if (-not $Force) {
                    $title = "Create Migration Batch"
                    $message = "Ready to create migration batch with the following mailboxes:`n"
                    $message += "- Ready: $($migrationState.ReadyMailboxes.Count)`n"
                    $message += "- Warning: $($migrationState.WarningMailboxes.Count)`n"
                    $message += "- Failed: $($migrationState.FailedMailboxes.Count)`n`n"
                    $message += "How would you like to proceed?"
                    
                    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Create batch with Ready mailboxes only"
                    $includeWarnings = New-Object System.Management.Automation.Host.ChoiceDescription "&Include Warnings", "Create batch with Ready and Warning mailboxes"
                    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel batch creation"
                    
                    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $includeWarnings, $no)
                    $result = $host.ui.PromptForChoice($title, $message, $options, 0)
                    
                    switch ($result) {
                        0 { 
                            $includeMailboxesWithWarnings = $false 
                            Write-Log -Message "User chose to create batch with Ready mailboxes only" -Level "INFO"
                            
                            # ShouldProcess check for the creation action
                            if (-not $PSCmdlet.ShouldProcess("Migration batch '$batchName'", "Create with $($migrationState.ReadyMailboxes.Count) ready mailboxes")) {
                                Write-Log -Message "Batch creation cancelled via ShouldProcess" -Level "WARNING"
                                $result.Success = $false
                                $result.Error = "Cancelled by user"
                                return $result
                            }
                        }
                        1 { 
                            $includeMailboxesWithWarnings = $true 
                            Write-Log -Message "User chose to include mailboxes with warnings" -Level "INFO"
                            
                            # ShouldProcess check for the creation action
                            $totalMailboxes = $migrationState.ReadyMailboxes.Count + $migrationState.WarningMailboxes.Count
                            if (-not $PSCmdlet.ShouldProcess("Migration batch '$batchName'", "Create with $totalMailboxes mailboxes (including warnings)")) {
                                Write-Log -Message "Batch creation cancelled via ShouldProcess" -Level "WARNING"
                                $result.Success = $false
                                $result.Error = "Cancelled by user"
                                return $result
                            }
                        }
                        2 { 
                            Write-Log -Message "User cancelled batch creation" -Level "WARNING"
                            $result.Success = $false
                            $result.Cancelled = $true
                            $result.Error = "Batch creation cancelled by user"
                            return $result
                        }
                    }
                }
                else {
                    # If Force is specified, assume we want to include mailboxes with warnings if IncludeWarnings is set
                    $includeMailboxesWithWarnings = $IncludeWarnings
                    Write-Log -Message "Force parameter specified, $($includeMailboxesWithWarnings ? 'including' : 'excluding') mailboxes with warnings" -Level "INFO"
                    
                    # ShouldProcess check even with -Force
                    $totalMailboxes = if ($includeMailboxesWithWarnings) {
                        $migrationState.ReadyMailboxes.Count + $migrationState.WarningMailboxes.Count
                    } else {
                        $migrationState.ReadyMailboxes.Count
                    }
                    
                    if (-not $PSCmdlet.ShouldProcess("Migration batch '$batchName'", "Create with $totalMailboxes mailboxes (Force specified)")) {
                        Write-Log -Message "Batch creation cancelled via ShouldProcess despite Force parameter" -Level "WARNING"
                        $result.Success = $false
                        $result.Error = "Cancelled by user"
                        return $result
                    }
                }
                
                $migrationState.CurrentStage = "CreatingBatch"
                Save-MigrationState -State $migrationState -Path $stateFilePath
            }
            
            # 6. Create migration batch
            if ($migrationState.CurrentStage -eq "CreatingBatch") {
                Write-Log -Message "Creating migration batch..." -Level "INFO"
                
                # Determine which mailboxes to include
                $mailboxesToMigrate = @()
                $mailboxesToMigrate += $migrationState.ReadyMailboxes
                $result.IncludedWarnings = $includeMailboxesWithWarnings
                
                if ($includeMailboxesWithWarnings) {
                    $mailboxesToMigrate += $migrationState.WarningMailboxes
                }
                
                if ($mailboxesToMigrate.Count -eq 0) {
                    throw "No mailboxes are eligible for migration"
                }
                
                # Get full mailbox objects
                $mailboxesForMigration = $validationResults | Where-Object { $mailboxesToMigrate -contains $_.EmailAddress }
                
                # Create the batch
                $batchResult = New-EXOMigrationBatch -Mailboxes $mailboxesForMigration `
                    -BatchName $batchName `
                    -TargetDeliveryDomain $script:Config.TargetDeliveryDomain `
                    -MigrationEndpointName $script:Config.MigrationEndpointName `
                    -CompleteAfterDays $script:Config.CompleteAfterDays `
                    -NotificationEmails $script:Config.NotificationEmails `
                    -UseBadItemLimits:$script:Config.UseBadItemLimitRecommendations
                
                if (-not $batchResult -or -not $batchResult.Identity) {
                    throw "Failed to create migration batch"
                }
                
                Write-Log -Message "Migration batch created: $($batchResult.Identity)" -Level "SUCCESS"
                Write-Log -Message "Batch status: $($batchResult.Status)" -Level "INFO"
                
                # Update state
                $migrationState.BatchCreated = $true
                $migrationState.BatchId = $batchResult.Identity
                $migrationState.CurrentStage = "Completed"
                $migrationState.CompletionTime = Get-Date
                Save-MigrationState -State $migrationState -Path $stateFilePath
                
                # Generate updated report with batch info
                $updatedReportPath = New-EXOMigrationReport -ValidationResults $validationResults `
                    -ReportPath $reportPath `
                    -BatchName $batchName `
                    -MigrationBatch $batchResult
                
                Write-Log -Message "Updated report generated: $updatedReportPath" -Level "INFO"
                
                # Send notification if configured
                if ($script:Config.NotificationEmails -and $script:Config.NotificationEmails.Count -gt 0) {
                    $mailSubject = "Migration Batch '$batchName' Created"
                    $mailBody = @"
<h2>Exchange Online Migration Batch Created</h2>
<p>A new migration batch has been created:</p>
<ul>
<li><strong>Batch Name:</strong> $batchName</li>
<li><strong>Batch ID:</strong> $($batchResult.Identity)</li>
<li><strong>Status:</strong> $($batchResult.Status)</li>
<li><strong>Mailboxes:</strong> $($mailboxesToMigrate.Count)</li>
<li><strong>Created:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</li>
</ul>
<p>Please see the attached report for details.</p>
"@
                    
                    Send-MigrationNotification -Subject $mailSubject -Body $mailBody -BodyAsHtml -Attachments $updatedReportPath
                }
                
                # Update result object
                $result.Success = $true
                $result.BatchCreated = $true
                $result.BatchId = $batchResult.Identity
                $result.BatchStatus = $batchResult.Status
                $result.MigratedMailboxCount = $mailboxesToMigrate.Count
                $result.ReadyMailboxes = $migrationState.ReadyMailboxes.Count
                $result.WarningMailboxes = $migrationState.WarningMailboxes.Count
                $result.FailedMailboxes = $migrationState.FailedMailboxes.Count
                $result.IncludedWarnings = $includeMailboxesWithWarnings
                $result.ReportPath = $updatedReportPath
                $result.ElapsedTime = (Get-Date) - $scriptStartTime
                
                # Record telemetry
                Record-MigrationMetric -MetricName "MigrationBatchCreated" -Value $batchName -Properties @{
                    TotalMailboxes = $mailboxesToMigrate.Count
                    ReadyMailboxes = $migrationState.ReadyMailboxes.Count
                    WarningMailboxes = $migrationState.WarningMailboxes.Count
                    FailedMailboxes = $migrationState.FailedMailboxes.Count
                    ElapsedSeconds = [math]::Round($result.ElapsedTime.TotalSeconds, 0)
                }
                
                return $result
            }
            
            # We shouldn't reach here if everything worked correctly
            throw "Migration process reached an unexpected state: $($migrationState.CurrentStage)"
        }
        catch {
            Write-Log -Message "Migration process failed: $_" -Level "ERROR"
            
            # Try to save the error state if possible
            if ($migrationState) {
                $migrationState.CurrentStage = "Failed"
                $migrationState.LastUpdated = Get-Date
                $migrationState | ConvertTo-Json -Depth 10 | Out-File -FilePath $stateFilePath -Force
            }
            
            # Update result object with error details
            $result.Success = $false
            $result.Error = $_.Exception.Message
            $result.ErrorDetails = $_
            $result.ElapsedTime = (Get-Date) - $scriptStartTime
            
            # Record error telemetry
            Record-MigrationMetric -MetricName "MigrationError" -Value $_.Exception.Message -Properties @{
                BatchName = $batchName
                ErrorType = $_.Exception.GetType().Name
                Stage = $migrationState.CurrentStage
                ElapsedSeconds = [math]::Round($result.ElapsedTime.TotalSeconds, 0)
            }
            
            return $result
        }
    }
    
    end {
        Write-Log -Message "Exchange Online Migration process completed. Elapsed time: $($result.ElapsedTime)" -Level "INFO"
    }
}