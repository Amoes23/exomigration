function New-EXOMigrationBatch {
    <#
    .SYNOPSIS
        Creates a new Exchange Online migration batch.
    
    .DESCRIPTION
        Creates a new migration batch in Exchange Online for the specified mailboxes.
        Supports customized batch parameters and intelligent BadItemLimit settings.
    
    .PARAMETER Mailboxes
        Array of mailbox email addresses to include in the migration batch.
    
    .PARAMETER BatchName
        Name for the migration batch. Must be unique.
    
    .PARAMETER TargetDeliveryDomain
        The mail.onmicrosoft.com domain for your tenant.
    
    .PARAMETER MigrationEndpointName
        Name of the migration endpoint to use for the migration.
    
    .PARAMETER CompleteAfterDays
        Number of days after which the migration batch will be automatically completed.
    
    .PARAMETER NotificationEmails
        Email addresses to notify about migration status.
    
    .PARAMETER StartAfterMinutes
        Minutes to wait before automatically starting the migration batch.
    
    .PARAMETER UseBadItemLimits
        When specified, uses recommended BadItemLimit values based on mailbox analysis.
    
    .PARAMETER ManualStart
        When specified, creates the batch but does not automatically start it.
    
    .EXAMPLE
        New-EXOMigrationBatch -Mailboxes $readyMailboxes -BatchName "Sales-Migration" -TargetDeliveryDomain "contoso.mail.onmicrosoft.com"
    
    .OUTPUTS
        The created migration batch object.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$Mailboxes,
        
        [Parameter(Mandatory = $true)]
        [string]$BatchName,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetDeliveryDomain,
        
        [Parameter(Mandatory = $true)]
        [string]$MigrationEndpointName,
        
        [Parameter(Mandatory = $false)]
        [int]$CompleteAfterDays = 1,
        
        [Parameter(Mandatory = $false)]
        [string[]]$NotificationEmails = @(),
        
        [Parameter(Mandatory = $false)]
        [int]$StartAfterMinutes = 15,
        
        [Parameter(Mandatory = $false)]
        [switch]$UseBadItemLimits,
        
        [Parameter(Mandatory = $false)]
        [switch]$ManualStart
    )
    
    try {
        if ($Mailboxes.Count -eq 0) {
            Write-Log -Message "No mailboxes provided for migration batch" -Level "ERROR" -ErrorCode "ERR017"
            throw "No mailboxes are ready for migration"
        }
        
        # Log the number of mailboxes
        Write-Log -Message "Preparing migration batch '$BatchName' for $($Mailboxes.Count) mailboxes..." -Level "INFO"
        
        # Check if a migration batch with the same name already exists
        $existingBatch = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
        if ($existingBatch) {
            Write-Log -Message "Migration batch with name '$BatchName' already exists with status: $($existingBatch.Status)" -Level "ERROR" -ErrorCode "ERR009"
            throw "Migration batch with name '$BatchName' already exists"
        }
        
        # Get migration endpoint
        $migrationEndpoint = Get-MigrationEndpoint -Identity $MigrationEndpointName -ErrorAction Stop
        if (-not $migrationEndpoint) {
            Write-Log -Message "Migration endpoint not found: $MigrationEndpointName" -Level "ERROR" -ErrorCode "ERR006"
            throw "Migration endpoint not found: $MigrationEndpointName"
        }
        
        # Set complete after date
        $completeAfterDate = (Get-Date).AddDays($CompleteAfterDays).ToUniversalTime()
        
        # Set start after time
        $startAfterDate = (Get-Date).AddMinutes($StartAfterMinutes)
        
        # If using BadItemLimit recommendations, we need to create move requests individually
        if ($UseBadItemLimits) {
            Write-Log -Message "Creating migration batch with custom BadItemLimit values..." -Level "INFO"
            
            # Create a new batch (empty) first
            $newBatchParams = @{
                Name = $BatchName
                SourceEndpoint = $migrationEndpoint.Identity
                TargetDeliveryDomain = $TargetDeliveryDomain
                CompleteAfter = $completeAfterDate
                StartAfter = $startAfterDate
                AutoStart = (-not $ManualStart)
            }
            
            if ($NotificationEmails -and $NotificationEmails.Count -gt 0) {
                $newBatchParams.Add('NotificationEmails', $NotificationEmails)
            }
            
            $newBatch = New-MigrationBatch @newBatchParams
            Write-Log -Message "Created base migration batch '$BatchName'" -Level "SUCCESS"
            
            # Now add mailboxes to the batch with individual BadItemLimit values
            $successCount = 0
            $errorCount = 0
            
            foreach ($mailbox in $Mailboxes) {
                try {
                    # If this is just an email address string
                    if ($mailbox -is [string]) {
                        $mailboxAddress = $mailbox
                        $badItemLimit = 10  # Default value
                    }
                    # If this is a result object from Test-EXOMailboxReadiness
                    elseif ($mailbox.EmailAddress -and $mailbox.PSObject.Properties.Name -contains "RecommendedBadItemLimit") {
                        $mailboxAddress = $mailbox.EmailAddress
                        $badItemLimit = if ($mailbox.RecommendedBadItemLimit -gt 0) {
                            $mailbox.RecommendedBadItemLimit
                        } else {
                            10  # Default value
                        }
                    }
                    # Otherwise, try to get the email address property
                    elseif ($mailbox.EmailAddress) {
                        $mailboxAddress = $mailbox.EmailAddress
                        $badItemLimit = 10  # Default value
                    }
                    else {
                        Write-Log -Message "Invalid mailbox object format. Skipping..." -Level "WARNING"
                        continue
                    }
                    
                    $moveRequestParams = @{
                        Identity = $mailboxAddress
                        BatchName = $BatchName
                        TargetDeliveryDomain = $TargetDeliveryDomain
                        BadItemLimit = $badItemLimit
                    }
                    
                    New-MoveRequest @moveRequestParams | Out-Null
                    $successCount++
                    
                    Write-Log -Message "Added mailbox $mailboxAddress to batch with BadItemLimit $badItemLimit" -Level "DEBUG"
                }
                catch {
                    $errorCount++
                    Write-Log -Message "Failed to add mailbox $mailboxAddress to batch: $_" -Level "ERROR"
                }
            }
            
            Write-Log -Message "Added $successCount of $($Mailboxes.Count) mailboxes to batch '$BatchName' (Errors: $errorCount)" -Level "INFO"
            
            # Start the batch if all went well and auto start is enabled
            if ($successCount -gt 0 -and -not $ManualStart) {
                Start-MigrationBatch -Identity $BatchName
                Write-Log -Message "Started migration batch '$BatchName'" -Level "SUCCESS"
            }
            else {
                Write-Log -Message "Migration batch '$BatchName' created but not started (Manual start requested or no mailboxes added successfully)" -Level "INFO"
            }
            
            # Refresh batch to get updated status
            $newBatch = Get-MigrationBatch -Identity $BatchName
        }
        else {
            # Standard approach - create CSV and use it for batch creation
            Write-Log -Message "Creating migration batch using standard CSV approach..." -Level "INFO"
            
            # Create a temporary CSV file with just the email addresses
            $tempCsvPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "$([Guid]::NewGuid().ToString())-$BatchName-ready.csv"
            
            try {
                # Create the CSV with explicit encoding
                "EmailAddress" | Out-File -FilePath $tempCsvPath -Encoding utf8
                
                # Add each mailbox address to the CSV
                foreach ($mailbox in $Mailboxes) {
                    if ($mailbox -is [string]) {
                        $mailbox | Out-File -FilePath $tempCsvPath -Append -Encoding utf8
                    }
                    elseif ($mailbox.EmailAddress) {
                        $mailbox.EmailAddress | Out-File -FilePath $tempCsvPath -Append -Encoding utf8
                    }
                }
                
                Write-Log -Message "Created temporary CSV for migration batch at: $tempCsvPath" -Level "DEBUG"
                
                # Create the migration batch
                Write-Log -Message "Creating migration batch '$BatchName'..." -Level "INFO"
                
                $newBatchParams = @{
                    Name = $BatchName
                    SourceEndpoint = $migrationEndpoint.Identity
                    TargetDeliveryDomain = $TargetDeliveryDomain
                    CSVData = [System.IO.File]::ReadAllBytes($tempCsvPath)
                    CompleteAfter = $completeAfterDate
                    StartAfter = $startAfterDate
                    AutoStart = (-not $ManualStart)
                }
                
                if ($NotificationEmails -and $NotificationEmails.Count -gt 0) {
                    $newBatchParams.Add('NotificationEmails', $NotificationEmails)
                }

                $newBatch = New-MigrationBatch @newBatchParams
                
                if ($newBatch -and -not $ManualStart) {
                    Write-Log -Message "Migration batch '$BatchName' created and will start automatically at $startAfterDate" -Level "SUCCESS"
                }
                else {
                    Write-Log -Message "Migration batch '$BatchName' created but not set to start automatically" -Level "SUCCESS"
                }
            }
            finally {
                # Clean up temporary CSV file
                if (Test-Path -Path $tempCsvPath) {
                    Remove-Item -Path $tempCsvPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        # Wait for batch to be created and retrieve status
        $timeoutMinutes = $script:Config.BatchCreationTimeoutMinutes ?? 10
        $endTime = (Get-Date).AddMinutes($timeoutMinutes)
        $batchCreated = $false
        
        Write-Log -Message "Waiting for migration batch to be created (timeout: $timeoutMinutes minutes)..." -Level "INFO"
        
        while ((Get-Date) -lt $endTime) {
            $batchStatus = Get-MigrationBatch -Identity $BatchName -ErrorAction SilentlyContinue
            
            if ($batchStatus) {
                $batchCreated = $true
                $newBatch = $batchStatus
                break
            }
            
            Write-Log -Message "Waiting for migration batch creation..." -Level "DEBUG"
            Start-Sleep -Seconds 5
        }
        
        if (-not $batchCreated) {
            Write-Log -Message "Migration batch creation timed out after $timeoutMinutes minutes" -Level "WARNING"
            Write-Log -Message "Check the Exchange Admin Center to verify if the batch was created" -Level "WARNING"
        }
        else {
            Write-Log -Message "Migration batch '$BatchName' created successfully with status: $($newBatch.Status)" -Level "SUCCESS"
            
            if (-not $ManualStart) {
                Write-Log -Message "Batch will start automatically in $StartAfterMinutes minutes" -Level "INFO"
            }
            
            Write-Log -Message "Batch will complete after: $completeAfterDate" -Level "INFO"
        }
        
        return $newBatch
    }
    catch {
        Write-Log -Message "Failed to create migration batch: $_" -Level "ERROR" -ErrorCode "ERR010"
        Write-Log -Message "Troubleshooting: Verify permissions and migration endpoint configuration" -Level "ERROR"
        throw
    }
}
