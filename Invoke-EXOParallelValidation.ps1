function Invoke-EXOParallelValidation {
    <#
    .SYNOPSIS
        Validates multiple mailboxes in parallel for migration readiness.
    
    .DESCRIPTION
        Processes mailboxes from a CSV file in parallel batches, using memory-efficient processing
        to handle large datasets. Results are serialized to disk to prevent excessive memory usage.
    
    .PARAMETER BatchFilePath
        Path to CSV file with mailboxes to validate. Must include an EmailAddress column.
    
    .PARAMETER ThrottleLimit
        Maximum number of mailboxes to process concurrently.
    
    .PARAMETER BatchSize
        Number of mailboxes to load into memory at once for large CSV files.
    
    .PARAMETER ValidationLevel
        Level of validation to perform: Basic, Standard, or Comprehensive.
    
    .PARAMETER IncludeInactiveMailboxes
        When specified, includes checking for mailbox inactivity.
    
    .PARAMETER OutputPath
        Directory to save validation results. If not specified, a temporary path is used.
    
    .EXAMPLE
        Invoke-EXOParallelValidation -BatchFilePath ".\Batches\Finance.csv" -ThrottleLimit 5
    
    .EXAMPLE
        Invoke-EXOParallelValidation -BatchFilePath ".\Batches\Sales.csv" -ValidationLevel Comprehensive -BatchSize 200
    
    .OUTPUTS
        [string] Path to the XML file containing the validation results.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$BatchFilePath,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 20)]
        [int]$ThrottleLimit = 5,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(10, 500)]
        [int]$BatchSize = 100,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Basic', 'Standard', 'Comprehensive')]
        [string]$ValidationLevel = 'Standard',
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeInactiveMailboxes,
        
        [Parameter(Mandatory = $false)]
        [ValidateScript({Test-Path $_ -PathType Container -IsValid})]
        [string]$OutputPath
    )
    
    try {
        Write-Log -Message "Starting parallel mailbox validation from file: $BatchFilePath" -Level "INFO"
        
        # Validate the batch file exists and has proper format
        if (-not (Test-Path -Path $BatchFilePath)) {
            throw "Batch file not found: $BatchFilePath"
        }
        
        # Check disk space before proceeding
        $diskSpaceOK = Test-DiskSpace -Paths @($OutputPath, (Split-Path -Path $BatchFilePath -Parent))
        if (-not $diskSpaceOK) {
            throw "Insufficient disk space for validation operation."
        }
        
        # Check if batch file contains the required header
        $csvHeaders = (Get-Content -Path $BatchFilePath -TotalCount 1).Split(',')
        if ($csvHeaders -notcontains "EmailAddress") {
            throw "CSV file must contain 'EmailAddress' column"
        }
        
        # Create output path for batch results if specified
        if ($OutputPath) {
            if (-not (Test-Path -Path $OutputPath)) {
                New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
                Write-Log -Message "Created output directory: $OutputPath" -Level "INFO"
            }
        }
        else {
            $tempDir = [System.IO.Path]::GetTempPath()
            $OutputPath = Join-Path -Path $tempDir -ChildPath "EXOMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
            Write-Log -Message "Created temporary output directory: $OutputPath" -Level "INFO"
        }
        
        # Import the CSV in batches to avoid loading entire file into memory
        $csv = New-Object System.IO.StreamReader $BatchFilePath
        
        # Skip header row
        $headerLine = $csv.ReadLine()
        
        # Count total mailboxes without loading all into memory
        $totalMailboxes = 0
        while (-not $csv.EndOfStream) {
            $null = $csv.ReadLine()
            $totalMailboxes++
        }
        $csv.Close()
        
        Write-Log -Message "Found $totalMailboxes mailboxes in CSV file" -Level "INFO"
        
        # Re-open file for processing
        $batchReader = New-Object System.IO.StreamReader $BatchFilePath
        
        # Skip header again for processing
        $null = $batchReader.ReadLine()
        
        # Process in batches
        $batchNumber = 1
        $resultsPath = Join-Path -Path $OutputPath -ChildPath "Combined_Results.xml"
        $processedCount = 0
        $startTime = Get-Date
        $results = New-Object System.Collections.ArrayList
        
        try {
            while (-not $batchReader.EndOfStream) {
                $currentBatch = @()
                $count = 0
                
                # Read current batch
                while (-not $batchReader.EndOfStream -and $count -lt $BatchSize) {
                    $line = $batchReader.ReadLine()
                    if (-not [string]::IsNullOrWhiteSpace($line)) {
                        $email = $line.Trim('"').Split(',')[0].Trim('"')  # Extract email from CSV line
                        if (-not [string]::IsNullOrWhiteSpace($email) -and $email -ne "EmailAddress") {
                            $currentBatch += @{ EmailAddress = $email }
                            $count++
                        }
                    }
                }
                
                if ($currentBatch.Count -eq 0) {
                    # Skip empty batches
                    continue
                }
                
                Write-Log -Message "Processing batch $batchNumber with $($currentBatch.Count) mailboxes..." -Level "INFO"
                Show-MigrationProgress -Activity "Validating mailboxes" -TotalItems $totalMailboxes -CompletedItems $processedCount -StartTime $startTime
                
                # Process current batch in parallel
                $batchResults = @()
                
                # Use the most appropriate parallel processing method for the current PowerShell version
                if ($PSVersionTable.PSVersion.Major -ge 7) {
                    # PowerShell 7+ parallel processing with ForEach-Object -Parallel
                    Write-Verbose "Using PowerShell 7+ parallel processing"
                    
                    $batchResults = $currentBatch | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
                        # Get parameters from parent scope
                        $emailAddress = $_.EmailAddress
                        $valLevel = $using:ValidationLevel
                        $checkInactive = $using:IncludeInactiveMailboxes.IsPresent
                        
                        # Process each mailbox
                        try {
                            $result = Test-EXOMailboxReadiness -EmailAddress $emailAddress `
                                -ValidationLevel $valLevel `
                                -IncludeInactiveMailboxes:$checkInactive
                            
                            # Return the validation result
                            return $result
                        }
                        catch {
                            # Create a failure result
                            return [PSCustomObject]@{
                                EmailAddress = $emailAddress
                                DisplayName = $emailAddress
                                Errors = @("Parallel processing failed: $($_.Exception.Message)")
                                ErrorCodes = @("ERR999")
                                Warnings = @()
                                OverallStatus = "Failed"
                                ValidationLevel = $valLevel
                            }
                        }
                    }
                }
                else {
                    # PowerShell 5.1 - Use runspaces
                    Write-Verbose "Using runspace pool for parallel processing"
                    
                    # Create runspace pool
                    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
                    $runspacePool.Open()
                    
                    $runspaces = @()
                    
                    # Create a runspace for each mailbox
                    foreach ($mailbox in $currentBatch) {
                        $powerShell = [powershell]::Create().AddScript({
                            param($EmailAddress, $ValidationLevel, $IncludeInactiveMailboxes)
                            
                            try {
                                # Call the test function for this mailbox
                                Test-EXOMailboxReadiness -EmailAddress $EmailAddress `
                                    -ValidationLevel $ValidationLevel `
                                    -IncludeInactiveMailboxes:$IncludeInactiveMailboxes
                            }
                            catch {
                                # Create a failure result
                                [PSCustomObject]@{
                                    EmailAddress = $EmailAddress
                                    DisplayName = $EmailAddress
                                    Errors = @("Runspace processing failed: $($_.Exception.Message)")
                                    ErrorCodes = @("ERR999")
                                    Warnings = @()
                                    OverallStatus = "Failed"
                                    ValidationLevel = $ValidationLevel
                                }
                            }
                        })
                        
                        $powerShell.AddParameter("EmailAddress", $mailbox.EmailAddress)
                        $powerShell.AddParameter("ValidationLevel", $ValidationLevel)
                        $powerShell.AddParameter("IncludeInactiveMailboxes", $IncludeInactiveMailboxes)
                        $powerShell.RunspacePool = $runspacePool
                        
                        $runspaces += [PSCustomObject]@{
                            PowerShell = $powerShell
                            Runspace = $powerShell.BeginInvoke()
                            EmailAddress = $mailbox.EmailAddress
                            Completed = $false
                        }
                    }
                    
                    # Collect results as they complete
                    do {
                        foreach ($runspace in $runspaces | Where-Object { $_.Runspace.IsCompleted -eq $true -and $_.Completed -eq $false }) {
                            $result = $runspace.PowerShell.EndInvoke($runspace.Runspace)
                            $runspace.PowerShell.Dispose()
                            $runspace.Completed = $true
                            
                            if ($null -ne $result) {
                                $batchResults += $result
                                
                                # Output status
                                $statusColor = switch ($result.OverallStatus) {
                                    "Ready" { "Green" }
                                    "Warning" { "Yellow" }
                                    "Failed" { "Red" }
                                    default { "White" }
                                }
                                
                                Write-Host "  - $($result.DisplayName) ($($result.EmailAddress)): " -NoNewline
                                Write-Host "$($result.OverallStatus)" -ForegroundColor $statusColor
                            }
                        }
                        
                        $completed = ($runspaces | Where-Object { $_.Completed -eq $true }).Count
                        $total = $runspaces.Count
                        
                        if ($completed -lt $total) {
                            Start-Sleep -Milliseconds 100
                        }
                    } while ($completed -lt $total)
                    
                    # Clean up
                    $runspacePool.Close()
                    $runspacePool.Dispose()
                }
                
                # Add batch results to overall results
                [void]$results.AddRange($batchResults)
                
                # Update progress
                $processedCount += $currentBatch.Count
                Show-MigrationProgress -Activity "Validating mailboxes" -TotalItems $totalMailboxes -CompletedItems $processedCount -StartTime $startTime
                
                # Save batch results to avoid keeping everything in memory
                $batchOutputPath = Join-Path -Path $OutputPath -ChildPath "Batch_${batchNumber}_Results.xml"
                $batchResults | Export-Clixml -Path $batchOutputPath -Force
                
                # Record metrics
                $readyCount = ($batchResults | Where-Object { $_.OverallStatus -eq "Ready" }).Count
                $warningCount = ($batchResults | Where-Object { $_.OverallStatus -eq "Warning" }).Count
                $failedCount = ($batchResults | Where-Object { $_.OverallStatus -eq "Failed" }).Count
                
                Record-MigrationMetric -MetricName "BatchValidation" -Value $batchNumber -Properties @{
                    BatchSize = $currentBatch.Count
                    Ready = $readyCount
                    Warning = $warningCount
                    Failed = $failedCount
                    ProcessingTime = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
                }
                
                # Increment batch counter
                $batchNumber++
                
                # Release memory
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }
        finally {
            # Close the reader
            $batchReader.Close()
            
            Show-MigrationProgress -Activity "Validating mailboxes" -Completed
        }
        
        # Combine all batch results into a single file
        Write-Log -Message "Combining batch results..." -Level "INFO"
        $combinedResults = @()
        
        Get-ChildItem -Path $OutputPath -Filter "Batch_*_Results.xml" | ForEach-Object {
            $batchResults = Import-Clixml -Path $_.FullName
            $combinedResults += $batchResults
            
            # Remove individual batch files to save space
            Remove-Item -Path $_.FullName -Force
        }
        
        # Export combined results
        $combinedResults | Export-Clixml -Path $resultsPath -Force
        
        # Final summary
        $finalReadyCount = ($combinedResults | Where-Object { $_.OverallStatus -eq "Ready" }).Count
        $finalWarningCount = ($combinedResults | Where-Object { $_.OverallStatus -eq "Warning" }).Count
        $finalFailedCount = ($combinedResults | Where-Object { $_.OverallStatus -eq "Failed" }).Count
        
        Write-Log -Message "Complete validation summary:" -Level "INFO"
        Write-Log -Message "  Ready: $finalReadyCount" -Level "INFO"
        Write-Log -Message "  Warning: $finalWarningCount" -Level "INFO"
        Write-Log -Message "  Failed: $finalFailedCount" -Level "INFO"
        Write-Log -Message "Completed validation of $($combinedResults.Count) mailboxes. Results saved to: $resultsPath" -Level "SUCCESS"
        
        # Record final metrics
        Record-MigrationMetric -MetricName "MailboxValidationComplete" -Value $combinedResults.Count -Properties @{
            Ready = $finalReadyCount
            Warning = $finalWarningCount
            Failed = $finalFailedCount
            TotalProcessingTime = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
            BatchesProcessed = $batchNumber - 1
        }
        
        return $resultsPath
    }
    catch {
        Write-Log -Message "Failed to perform parallel mailbox validation: $_" -Level "ERROR"
        throw
    }
}