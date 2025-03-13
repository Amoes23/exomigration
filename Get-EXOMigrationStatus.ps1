function Get-EXOMigrationStatus {
    <#
    .SYNOPSIS
        Gets the status of Exchange Online migration batches.
    
    .DESCRIPTION
        Retrieves detailed information about migration batches in Exchange Online,
        including progress, statistics, and error details. Can be used to monitor
        ongoing migrations initiated with Start-EXOMigration.
    
    .PARAMETER BatchName
        Optional name of a specific migration batch to check. If not specified,
        returns status for all migration batches.
    
    .PARAMETER IncludeInactive
        When specified, includes completed and failed batches in the results.
        By default, only active batches are returned.
    
    .PARAMETER DetailLevel
        Level of detail to include in the results:
        - Basic: Basic batch information only
        - Detailed: Includes mailbox statistics
        - Full: Includes mailbox statistics and error details (default)
    
    .PARAMETER LastDays
        Number of days to look back for migration batches. Default is 7 days.
    
    .PARAMETER IncludeReport
        When specified, includes the full diagnostic report for each mailbox move.
        This significantly increases the amount of data returned.
    
    .EXAMPLE
        Get-EXOMigrationStatus
    
    .EXAMPLE
        Get-EXOMigrationStatus -BatchName "Finance-Migration" -DetailLevel Full
    
    .EXAMPLE
        Get-EXOMigrationStatus -IncludeInactive -LastDays 30
    
    .OUTPUTS
        [PSCustomObject[]] Collection of migration batch status objects.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$BatchName,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeInactive,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Basic', 'Detailed', 'Full')]
        [string]$DetailLevel = 'Full',
        
        [Parameter(Mandatory = $false)]
        [int]$LastDays = 7,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeReport
    )
    
    try {
        Write-Log -Message "Retrieving migration batch status..." -Level "INFO"
        
        # Check if connected to Exchange Online
        try {
            $null = Get-MigrationConfig -ErrorAction Stop
        }
        catch {
            Write-Log -Message "Not connected to Exchange Online. Connecting..." -Level "WARNING"
            $connected = Connect-EXOMigrationServices
            if (-not $connected) {
                throw "Failed to connect to Exchange Online. Please run Connect-ExchangeOnline first."
            }
        }
        
        # Get migration batches
        $filter = if ($BatchName) {
            { $_.Identity -eq $BatchName }
        }
        elseif (-not $IncludeInactive) {
            { $_.Status -notlike "*Completed*" -and $_.Status -ne "Failed" }
        }
        else {
            { $_.CreationTime -ge (Get-Date).AddDays(-$LastDays) }
        }
        
        $migrationBatches = Get-MigrationBatch | Where-Object $filter
        
        if (-not $migrationBatches) {
            if ($BatchName) {
                Write-Log -Message "No migration batch found with name: $BatchName" -Level "WARNING"
            }
            else {
                Write-Log -Message "No active migration batches found" -Level "INFO"
            }
            return @()
        }
        
        $results = @()
        
        foreach ($batch in $migrationBatches) {
            $batchResult = [PSCustomObject]@{
                BatchName = $batch.Identity
                Status = $batch.Status
                CreationTime = $batch.CreationTime
                StartTime = $batch.StartTime
                CompleteAfter = $batch.CompleteAfter
                MailboxCount = $batch.TotalCount
                TargetDeliveryDomain = $batch.TargetDeliveryDomain
                NotificationEmails = $batch.NotificationEmails -join ', '
                ActiveMailboxes = 0
                CompletedMailboxes = 0
                FailedMailboxes = 0
                SyncedMailboxes = 0
                FinalizedMailboxes = 0
                PercentComplete = 0
                HasErrors = $false
                MailboxDetails = @()
                LastUpdated = Get-Date
            }
            
            # Get mailbox statistics for the batch
            if ($DetailLevel -ne 'Basic') {
                $moveRequests = Get-MoveRequest -BatchName $batch.Identity
                
                $batchResult.ActiveMailboxes = ($moveRequests | Where-Object { $_.Status -eq "InProgress" }).Count
                $batchResult.CompletedMailboxes = ($moveRequests | Where-Object { $_.Status -eq "Completed" }).Count
                $batchResult.FailedMailboxes = ($moveRequests | Where-Object { $_.Status -eq "Failed" }).Count
                $batchResult.SyncedMailboxes = ($moveRequests | Where-Object { 
                    $_.Status -eq "Completed" -or 
                    ($_.Status -eq "InProgress" -and $_.PercentComplete -ge 95) 
                }).Count
                $batchResult.FinalizedMailboxes = ($moveRequests | Where-Object { $_.Status -eq "Completed" -and $_.ArchiveStatus -ne "Pending" }).Count
                
                # Calculate overall percentage
                if ($batch.TotalCount -gt 0) {
                    $totalProgress = ($moveRequests | Measure-Object -Property PercentComplete -Average).Average
                    $batchResult.PercentComplete = [math]::Round($totalProgress, 2)
                }
                
                # Add detailed mailbox information for Full detail level
                if ($DetailLevel -eq 'Full') {
                    $detailedRequests = @()
                    
                    foreach ($request in $moveRequests) {
                        try {
                            $statsParams = @{
                                Identity = $request.Identity
                                ErrorAction = 'Stop'
                            }
                            
                            if ($IncludeReport) {
                                $statsParams.Add('IncludeReport', $true)
                                $statsParams.Add('DiagnosticInfo', 'Verbose')
                            }
                            
                            $moveStats = Get-MoveRequestStatistics @statsParams
                            
                            $requestDetails = [PSCustomObject]@{
                                Identity = $request.Identity
                                Status = $request.Status
                                PercentComplete = $request.PercentComplete
                                TotalItemCount = $moveStats.TotalItemCount
                                TotalItemSize = $moveStats.TotalItemSize
                                ItemsTransferred = $moveStats.ItemsTransferred
                                BytesTransferred = $moveStats.BytesTransferred
                                BadItemsEncountered = $moveStats.BadItemsEncountered
                                BadItemLimit = $moveStats.BadItemLimit
                                LargeItemsEncountered = $moveStats.LargeItemsEncountered
                                LargeItemLimit = $moveStats.LargeItemLimit
                                QueuedTimestamp = $moveStats.QueuedTimestamp
                                StartTimestamp = $moveStats.StartTimestamp
                                LastUpdateTimestamp = $moveStats.LastUpdateTimestamp
                                FailureCode = $moveStats.FailureCode
                                FailureSeverity = $moveStats.FailureSeverity
                                ErrorSummary = $moveStats.Message
                                HasErrors = $false
                                ErrorDetails = @()
                            }
                            
                            # Add error information if available
                            if ($moveStats.Report -and $moveStats.Report.Entries) {
                                $errors = $moveStats.Report.Entries | Where-Object { $_.Message -like "*error*" -or $_.Message -like "*failed*" }
                                
                                if ($errors) {
                                    $requestDetails.HasErrors = $true
                                    $batchResult.HasErrors = $true
                                    
                                    foreach ($error in $errors) {
                                        $requestDetails.ErrorDetails += [PSCustomObject]@{
                                            Timestamp = $error.CreationTime
                                            Message = $error.Message
                                            Type = $error.Type
                                        }
                                    }
                                }
                            }
                            
                            $detailedRequests += $requestDetails
                        }
                        catch {
                            Write-Log -Message "Failed to get detailed statistics for $($request.Identity): $_" -Level "WARNING"
                            
                            # Add basic information without details
                            $detailedRequests += [PSCustomObject]@{
                                Identity = $request.Identity
                                Status = $request.Status
                                PercentComplete = $request.PercentComplete
                                ErrorSummary = "Failed to retrieve detailed statistics: $_"
                                HasErrors = $true
                            }
                        }
                    }
                    
                    $batchResult.MailboxDetails = $detailedRequests
                }
            }
            
            $results += $batchResult
        }
        
        Write-Log -Message "Retrieved status for $($results.Count) migration batches" -Level "SUCCESS"
        return $results
    }
    catch {
        Write-Log -Message "Failed to retrieve migration status: $_" -Level "ERROR"
        throw
    }
}
