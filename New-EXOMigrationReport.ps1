function New-EXOMigrationReport {
    <#
    .SYNOPSIS
        Generates an HTML report for Exchange Online migration readiness.
    
    .DESCRIPTION
        Creates a comprehensive HTML report based on mailbox validation results.
        The report includes overall statistics, detailed mailbox information,
        migration guidance, and actionable recommendations.
    
    .PARAMETER ValidationResults
        Collection of validation results from Test-EXOMailboxReadiness.
    
    .PARAMETER ReportPath
        Path where the report should be saved. If not specified, uses the default
        from configuration.
    
    .PARAMETER BatchName
        Name of the migration batch for the report.
    
    .PARAMETER MigrationBatch
        Optional migration batch object to include details in the report.
    
    .PARAMETER TemplatePath
        Path to the HTML template file. If not specified, uses the default template.
    
    .PARAMETER OpenReport
        When specified, automatically opens the report in the default browser after creation.
    
    .EXAMPLE
        New-EXOMigrationReport -ValidationResults $results -BatchName "Finance-Migration"
    
    .EXAMPLE
        New-EXOMigrationReport -ValidationResults $results -BatchName "Sales-Migration" -MigrationBatch $batchObj -OpenReport
    
    .OUTPUTS
        [string] Path to the generated HTML report.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$ValidationResults,
        
        [Parameter(Mandatory = $false)]
        [string]$ReportPath,
        
        [Parameter(Mandatory = $true)]
        [string]$BatchName,
        
        [Parameter(Mandatory = $false)]
        [object]$MigrationBatch,
        
        [Parameter(Mandatory = $false)]
        [string]$TemplatePath,
        
        [Parameter(Mandatory = $false)]
        [switch]$OpenReport
    )
    
    try {
        Write-Log -Message "Generating migration report for batch: $BatchName" -Level "INFO"
        
        # Use config template path if not specified in parameters
        if (-not $TemplatePath -and $script:Config -and $script:Config.HTMLTemplatePath) {
            $TemplatePath = $script:Config.HTMLTemplatePath
            
            # Validate template path
            if (-not (Test-Path -Path $TemplatePath)) {
                Write-Log -Message "HTML template not found at configured path: $TemplatePath" -Level "WARNING"
                $TemplatePath = $null
            }
        }
        
        # Use default report path from config if not specified
        if (-not $ReportPath -and $script:Config -and $script:Config.ReportPath) {
            $defaultReportDir = $script:Config.ReportPath
            
            # Create report directory if it doesn't exist
            if (-not (Test-Path -Path $defaultReportDir)) {
                New-Item -ItemType Directory -Path $defaultReportDir -Force | Out-Null
                Write-Log -Message "Created report directory: $defaultReportDir" -Level "INFO"
            }
            
            # Generate report filename with timestamp
            $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
            $reportFilename = "$BatchName-report-$timestamp.html"
            $ReportPath = Join-Path -Path $defaultReportDir -ChildPath $reportFilename
        }
        elseif (-not $ReportPath) {
            # If no report path specified or in config, use current directory
            $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
            $reportFilename = "$BatchName-report-$timestamp.html"
            $ReportPath = Join-Path -Path (Get-Location).Path -ChildPath $reportFilename
        }
        
        # Ensure report path has directory
        $reportDir = Split-Path -Path $ReportPath -Parent
        if (-not (Test-Path -Path $reportDir)) {
            New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
            Write-Log -Message "Created report directory: $reportDir" -Level "INFO"
        }
        
        # Generate the HTML report
        $reportFile = Export-HTMLReport -TestResults $ValidationResults `
            -TemplatePath $TemplatePath `
            -ReportPath $ReportPath `
            -BatchName $BatchName `
            -MigrationBatch $MigrationBatch
        
        if (-not $reportFile -or -not (Test-Path -Path $reportFile)) {
            Write-Log -Message "Failed to generate HTML report" -Level "ERROR" -ErrorCode "ERR018"
            throw "Failed to generate HTML report"
        }
        
        Write-Log -Message "HTML report generated: $reportFile" -Level "SUCCESS"
        
        # Open the report if requested
        if ($OpenReport) {
            try {
                Write-Log -Message "Opening report in default browser" -Level "INFO"
                Start-Process -FilePath $reportFile
            }
            catch {
                Write-Log -Message "Failed to open report: $_" -Level "WARNING"
                Write-Log -Message "You can manually open the report at: $reportFile" -Level "INFO"
            }
        }
        
        return $reportFile
    }
    catch {
        Write-Log -Message "Failed to generate migration report: $_" -Level "ERROR"
        throw
    }
}
