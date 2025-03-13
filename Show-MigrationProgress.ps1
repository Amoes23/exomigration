function Show-MigrationProgress {
    <#
    .SYNOPSIS
        Displays progress information for migration operations.
    
    .DESCRIPTION
        Provides enhanced progress display for migration operations with ETA calculation.
    
    .PARAMETER Activity
        The name of the activity for which progress is being reported.
    
    .PARAMETER TotalItems
        The total number of items to be processed.
    
    .PARAMETER CompletedItems
        The number of items that have been processed so far.
    
    .PARAMETER StartTime
        The time when processing started, used for ETA calculation.
    
    .PARAMETER Completed
        When specified, marks the operation as completed and clears the progress bar.
    
    .EXAMPLE
        Show-MigrationProgress -Activity "Validating Mailboxes" -TotalItems 100 -CompletedItems 25 -StartTime $startTime
    
    .OUTPUTS
        None. Progress is displayed to the user.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Activity,
        
        [Parameter(Mandatory = $true)]
        [int]$TotalItems,
        
        [Parameter(Mandatory = $true)]
        [int]$CompletedItems,
        
        [Parameter(Mandatory = $false)]
        [DateTime]$StartTime = $null,
        
        [Parameter(Mandatory = $false)]
        [switch]$Completed
    )
    
    if ($Completed) {
        Write-Progress -Activity $Activity -Completed
        return
    }
    
    $percentComplete = [math]::Min(100, [math]::Round(($CompletedItems / $TotalItems) * 100, 0))
    
    # Calculate ETA if start time was provided
    $etaString = ""
    if ($StartTime -and $CompletedItems -gt 0) {
        $elapsed = (Get-Date) - $StartTime
        $itemsRemaining = $TotalItems - $CompletedItems
        
        if ($elapsed.TotalSeconds -gt 0) {
            $itemsPerSecond = $CompletedItems / $elapsed.TotalSeconds
            if ($itemsPerSecond -gt 0) {
                $secondsRemaining = $itemsRemaining / $itemsPerSecond
                $eta = (Get-Date).AddSeconds($secondsRemaining)
                $etaString = " | ETA: $(Get-Date $eta -Format 'HH:mm:ss')"
            }
        }
    }
    
    $statusMessage = "Processed $CompletedItems of $TotalItems items ($percentComplete%)$etaString"
    Write-Progress -Activity $Activity -Status $statusMessage -PercentComplete $percentComplete
}