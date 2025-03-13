function Record-MigrationMetric {
    <#
    .SYNOPSIS
        Records metrics about the migration process.
    
    .DESCRIPTION
        Saves performance and operational metrics about the migration process for
        analysis and monitoring. Metrics are stored in a JSON file for easy parsing.
    
    .PARAMETER MetricName
        The name of the metric being recorded.
    
    .PARAMETER Value
        The value of the metric being recorded.
    
    .PARAMETER Properties
        Additional properties to store with the metric (optional).
    
    .PARAMETER Append
        When specified, appends to existing metrics file instead of overwriting.
        Default is to append.
    
    .EXAMPLE
        Record-MigrationMetric -MetricName "MailboxValidationTime" -Value 120.5 -Properties @{Mailbox="user@contoso.com"; Size="2.5 GB"}
    
    .OUTPUTS
        None. Metrics are written to a file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MetricName,
        
        [Parameter(Mandatory = $true)]
        [object]$Value,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Properties = @{},
        
        [Parameter(Mandatory = $false)]
        [switch]$Append = $true
    )
    
    try {
        $metric = [PSCustomObject]@{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Name = $MetricName
            Value = $Value
            Properties = $Properties
        }
        
        # Ensure log path exists
        if (-not (Test-Path -Path $script:Config.LogPath)) {
            New-Item -ItemType Directory -Path $script:Config.LogPath -Force | Out-Null
        }
        
        $metricsFile = Join-Path $script:Config.LogPath "MigrationMetrics.json"
        
        # Load existing metrics if appending
        if ($Append -and (Test-Path -Path $metricsFile)) {
            try {
                $existingMetrics = Get-Content -Path $metricsFile -Raw | ConvertFrom-Json -ErrorAction Stop
                # Convert to array if it's a single object
                if ($existingMetrics -isnot [array]) {
                    $existingMetrics = @($existingMetrics)
                }
                $metrics = $existingMetrics + $metric
            }
            catch {
                Write-Verbose "Could not read existing metrics, starting new file: $_"
                $metrics = @($metric)
            }
        }
        else {
            $metrics = @($metric)
        }
        
        # Save metrics
        $metrics | ConvertTo-Json -Depth 4 | Out-File -FilePath $metricsFile -Force
        
        Write-Verbose "Recorded metric: $MetricName = $Value"
    }
    catch {
        # Don't let metrics recording failures break the main process
        Write-Verbose "Failed to record metric: $_"
    }
}