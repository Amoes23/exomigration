function Test-DiskSpace {
    <#
    .SYNOPSIS
        Checks if there is enough disk space available for the operation.
    
    .DESCRIPTION
        Verifies that the specified paths have enough disk space available
        before proceeding with operations that may require significant space.
        This helps prevent failures due to insufficient disk space during
        migration operations.
    
    .PARAMETER Paths
        An array of paths to check for disk space.
    
    .PARAMETER ThresholdMB
        Minimum required disk space in MB. Defaults to value from configuration
        or 500 MB if not specified.
    
    .PARAMETER ErrorIfBelowThreshold
        When specified, throws an error if available space is below threshold
        instead of just returning false.
    
    .EXAMPLE
        Test-DiskSpace -Paths @("C:\Logs", "C:\Reports") -ThresholdMB 500
    
    .EXAMPLE
        Test-DiskSpace -Paths "C:\Migration" -ErrorIfBelowThreshold
    
    .OUTPUTS
        [bool] True if enough disk space is available, False otherwise.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Paths,
        
        [Parameter(Mandatory = $false)]
        [int]$ThresholdMB = 0,
        
        [Parameter(Mandatory = $false)]
        [switch]$ErrorIfBelowThreshold
    )
    
    try {
        # Use config value if not specified directly
        if ($ThresholdMB -eq 0 -and $script:Config -and $script:Config.DiskSpaceThresholdMB) {
            $ThresholdMB = $script:Config.DiskSpaceThresholdMB
        }
        elseif ($ThresholdMB -eq 0) {
            $ThresholdMB = 500  # Default value if not in config
        }
        
        Write-Verbose "Checking for at least $ThresholdMB MB free disk space in specified paths"
        $allPathsHaveSpace = $true
        $lowSpacePaths = @()
        
        foreach ($path in $Paths) {
            # Get drive root for the path
            $drive = $null
            
            # Try to resolve path
            try {
                $resolvedPath = Resolve-Path $path -ErrorAction SilentlyContinue
                if ($resolvedPath) {
                    $drive = (Split-Path -Qualifier $resolvedPath) + "\"
                }
            }
            catch {
                # Path might not exist yet
            }
            
            if (-not $drive) {
                # If path doesn't exist yet, get parent folder's drive
                $parentPath = Split-Path -Parent $path
                while ($parentPath -and -not (Test-Path $parentPath)) {
                    $parentPath = Split-Path -Parent $parentPath
                }
                
                if ($parentPath) {
                    $drive = (Resolve-Path $parentPath | Split-Path -Qualifier) + "\"
                }
                else {
                    # If still can't resolve, use current location
                    $drive = (Get-Location | Split-Path -Qualifier) + "\"
                }
            }
            
            # Get drive free space
            try {
                $driveInfo = Get-PSDrive -Name $drive[0] -PSProvider FileSystem
                $freeSpaceMB = [math]::Round($driveInfo.Free / 1MB, 2)
                
                Write-Verbose "Drive $drive has $freeSpaceMB MB free space"
                
                if ($freeSpaceMB -lt $ThresholdMB) {
                    $allPathsHaveSpace = $false
                    $lowSpacePaths += [PSCustomObject]@{
                        Path = $path
                        Drive = $drive
                        FreeSpaceMB = $freeSpaceMB
                        RequiredMB = $ThresholdMB
                    }
                    
                    Write-Log -Message "Insufficient disk space on $drive for $path. Available: $freeSpaceMB MB, Required: $ThresholdMB MB" -Level "WARNING"
                }
            }
            catch {
                Write-Log -Message "Error checking disk space for drive $drive`: $_" -Level "WARNING"
                $allPathsHaveSpace = $false
                $lowSpacePaths += [PSCustomObject]@{
                    Path = $path
                    Drive = $drive
                    FreeSpaceMB = "Unknown"
                    RequiredMB = $ThresholdMB
                    Error = $_.Exception.Message
                }
            }
        }
        
        if (-not $allPathsHaveSpace) {
            $message = "Insufficient disk space detected for operation:"
            foreach ($item in $lowSpacePaths) {
                if ($item.FreeSpaceMB -eq "Unknown") {
                    $message += "`n  - Path: $($item.Path) on drive $($item.Drive): Error checking free space - $($item.Error)"
                }
                else {
                    $message += "`n  - Path: $($item.Path) on drive $($item.Drive): $($item.FreeSpaceMB) MB available, $($item.RequiredMB) MB required"
                }
            }
            
            if ($ErrorIfBelowThreshold) {
                Write-Log -Message $message -Level "ERROR"
                throw $message
            }
            else {
                Write-Log -Message $message -Level "WARNING"
            }
        }
        else {
            Write-Verbose "All paths have sufficient disk space"
        }
        
        return $allPathsHaveSpace
    }
    catch {
        $errorMessage = "Error checking disk space: $_"
        Write-Log -Message $errorMessage -Level "ERROR"
        
        if ($ErrorIfBelowThreshold) {
            throw $errorMessage
        }
        
        return $false
    }
}
