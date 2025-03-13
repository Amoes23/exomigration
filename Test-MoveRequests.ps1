function Test-MoveRequests {
    <#
    .SYNOPSIS
        Tests for existing move requests that might affect migration.
    
    .DESCRIPTION
        Checks for pending move requests on the mailbox that might interfere with
        a new migration attempt. Identifies orphaned move requests and provides 
        options to remove them if necessary.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .PARAMETER RemoveAndResubmitMoveRequests
        When specified, attempts to remove existing move requests to allow a new migration.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-MoveRequests -EmailAddress "user@contoso.com" -Results $results
    
    .EXAMPLE
        Test-MoveRequests -EmailAddress "user@contoso.com" -Results $results -RemoveAndResubmitMoveRequests
    
    .OUTPUTS
        [bool] Returns $true if the test was successful (even if issues were found), $false if the test failed.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results,
        
        [Parameter(Mandatory = $false)]
        [switch]$RemoveAndResubmitMoveRequests
    )
    
    try {
        # Check for pending move requests
        try {
            $moveRequest = Get-MoveRequest -Identity $EmailAddress -ErrorAction Stop
            
            $Results.PendingMoveRequest = $true
            $Results.MoveRequestStatus = $moveRequest.Status
            
            # Get detailed move request information
            try {
                $moveRequestDetails = Get-MoveRequestStatistics -Identity $EmailAddress -IncludeReport -ErrorAction Stop
                $Results | Add-Member -NotePropertyName "MoveRequestDetails" -NotePropertyValue $moveRequestDetails -Force
            }
            catch [Microsoft.Exchange.Management.Migration.MoveRequestNotFoundException] {
                Write-Log -Message "Move request statistics not found for $EmailAddress, but move request exists" -Level "WARNING"
                # Not critical - continue without detailed stats
            }
            catch {
                Write-Log -Message "Failed to get move request statistics: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
                # Not critical - continue without detailed stats
            }
            
            $Results.Errors += "Pending move request found with status: $($moveRequest.Status)"
            $Results.ErrorCodes += "ERR016"
            Write-Log -Message "Error: Pending move request found with status: $($moveRequest.Status) for $EmailAddress" -Level "ERROR" -ErrorCode "ERR016"
            Write-Log -Message "Troubleshooting: Wait for the existing move request to complete or remove it with Remove-MoveRequest" -Level "INFO"
            
            # Handle remove and resubmit option
            if ($RemoveAndResubmitMoveRequests) {
                try {
                    Write-Log -Message "Removing existing move request for $EmailAddress" -Level "INFO"
                    Remove-MoveRequest -Identity $EmailAddress -Confirm:$false -ErrorAction Stop
                    Write-Log -Message "Existing move request removed successfully" -Level "SUCCESS"
                    
                    # Update results to reflect removal
                    $Results.PendingMoveRequest = $false
                    $Results.MoveRequestStatus = "Removed"
                    $Results.Errors = $Results.Errors | Where-Object { $_ -notlike "*Pending move request found*" }
                    $Results.ErrorCodes = $Results.ErrorCodes | Where-Object { $_ -ne "ERR016" }
                    
                    # Add a warning instead
                    $Results.Warnings += "Previous move request was removed and will be resubmitted"
                    Write-Log -Message "Mailbox will be included in new migration batch" -Level "INFO"
                }
                catch [Microsoft.Exchange.Management.Migration.MoveRequestNotFoundException] {
                    Write-Log -Message "Move request no longer exists for $EmailAddress" -Level "WARNING"
                    $Results.Warnings += "Move request no longer exists - may have been removed by another process"
                    
                    # Update results to reflect removal
                    $Results.PendingMoveRequest = $false
                    $Results.MoveRequestStatus = "NotFound"
                }
                catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
                    Write-Log -Message "Insufficient permissions to remove move request for $EmailAddress" -Level "ERROR"
                    $Results.Warnings += "Failed to remove existing move request due to insufficient permissions"
                }
                catch {
                    Write-Log -Message "Failed to remove move request for $EmailAddress`: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
                    $Results.Warnings += "Failed to remove existing move request: $($_.Exception.Message)"
                }
            }
        }
        catch [Microsoft.Exchange.Management.Migration.MoveRequestNotFoundException] {
            # No move request found, this is good
            $Results.PendingMoveRequest = $false
            $Results.MoveRequestStatus = "None"
        }
        catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
            Write-Log -Message "Insufficient permissions to check move requests for $EmailAddress" -Level "ERROR"
            $Results.Warnings += "Insufficient permissions to check move requests. Verify your Exchange permissions."
            return $false
        }
        catch {
            Write-Log -Message "Error checking move requests: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
            $Results.Warnings += "Error checking move requests: $($_.Exception.Message)"
            return $false
        }
        
        # Check for orphaned move requests
        try {
            $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
            
            try {
                $orphanedMoveRequests = Get-MoveRequest | Where-Object { 
                    $_.DisplayName -eq $mailbox.DisplayName -and $_.Identity -ne $EmailAddress 
                }
                
                if ($orphanedMoveRequests) {
                    foreach ($orphanedMove in $orphanedMoveRequests) {
                        $Results.Errors += "Orphaned move request found: $($orphanedMove.Identity) with status $($orphanedMove.Status)"
                        $Results.ErrorCodes += "ERR016"
                        
                        Write-Log -Message "Error: Orphaned move request found: $($orphanedMove.Identity) with status $($orphanedMove.Status) for $EmailAddress" -Level "ERROR" -ErrorCode "ERR016"
                        Write-Log -Message "Troubleshooting: Remove orphaned move request with: Remove-MoveRequest -Identity '$($orphanedMove.Identity)'" -Level "INFO"
                        
                        # Handle remove and resubmit option for orphaned requests
                        if ($RemoveAndResubmitMoveRequests) {
                            try {
                                Write-Log -Message "Removing orphaned move request: $($orphanedMove.Identity)" -Level "INFO"
                                Remove-MoveRequest -Identity $orphanedMove.Identity -Confirm:$false -ErrorAction Stop
                                Write-Log -Message "Orphaned move request removed successfully" -Level "SUCCESS"
                                
                                # Add a warning
                                $Results.Warnings += "Orphaned move request was removed: $($orphanedMove.Identity)"
                            }
                            catch [Microsoft.Exchange.Management.Migration.MoveRequestNotFoundException] {
                                Write-Log -Message "Orphaned move request no longer exists: $($orphanedMove.Identity)" -Level "WARNING"
                                $Results.Warnings += "Orphaned move request no longer exists - may have been removed by another process"
                            }
                            catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
                                Write-Log -Message "Insufficient permissions to remove orphaned move request: $($orphanedMove.Identity)" -Level "ERROR"
                                $Results.Warnings += "Failed to remove orphaned move request due to insufficient permissions"
                            }
                            catch {
                                Write-Log -Message "Failed to remove orphaned move request $($orphanedMove.Identity): $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
                                $Results.Warnings += "Failed to remove orphaned move request: $($_.Exception.Message)"
                            }
                        }
                    }
                }
            }
            catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
                Write-Log -Message "Insufficient permissions to search for orphaned move requests" -Level "WARNING"
                $Results.Warnings += "Insufficient permissions to search for orphaned move requests"
            }
            catch {
                Write-Log -Message "Error searching for orphaned move requests: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
                $Results.Warnings += "Error searching for orphaned move requests: $($_.Exception.Message)"
            }
        }
        catch [Microsoft.Exchange.Management.Tasks.ManagementObjectNotFoundException] {
            Write-Log -Message "Mailbox not found: $EmailAddress" -Level "ERROR"
            $Results.Errors += "Mailbox not found in Exchange Online"
            return $false
        }
        catch [Microsoft.Exchange.Management.RbacInsufficientAccessException] {
            Write-Log -Message "Insufficient permissions to get mailbox: $EmailAddress" -Level "ERROR"
            $Results.Warnings += "Insufficient permissions to get mailbox information"
            return $false
        }
        catch {
            Write-Log -Message "Error getting mailbox: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
            $Results.Warnings += "Error getting mailbox: $($_.Exception.Message)"
            return $false
        }
        
        # Check for incomplete migrations
        if ($script:Config.CheckIncompleteMailboxMoves) {
            # Get migration history that might indicate previous issues
            try {
                $moveHistory = Get-MoveRequestStatistics -Identity $EmailAddress -IncludeReport -ErrorAction SilentlyContinue
                
                if ($moveHistory) {
                    # Look for signs of problems in previous migrations
                    $hasFailures = ($moveHistory.Report.Entries | Where-Object { $_.Message -like "*fail*" -or $_.Message -like "*error*" }).Count -gt 0
                    $hasBadItems = $moveHistory.BadItemsEncountered -gt 0
                    $hasLargeItems = $moveHistory.LargeItemsEncountered -gt 0
                    
                    if ($hasFailures -or $hasBadItems -or $hasLargeItems) {
                        $Results | Add-Member -NotePropertyName "HasIncompleteMoves" -NotePropertyValue $true -Force
                        $Results | Add-Member -NotePropertyName "IncompleteMoveReason" -NotePropertyValue @() -Force
                        
                        if ($hasFailures) {
                            $Results.IncompleteMoveReason += "Previous migration attempts had failures"
                            $Results.Warnings += "Previous migration attempts had failures - check move request report for details"
                            Write-Log -Message "Warning: Previous migration attempts for $EmailAddress had failures" -Level "WARNING"
                        }
                        
                        if ($hasBadItems) {
                            $Results.IncompleteMoveReason += "Previous migration encountered $($moveHistory.BadItemsEncountered) bad items"
                            $Results.Warnings += "Previous migration encountered $($moveHistory.BadItemsEncountered) bad items"
                            Write-Log -Message "Warning: Previous migration for $EmailAddress encountered $($moveHistory.BadItemsEncountered) bad items" -Level "WARNING"
                        }
                        
                        if ($hasLargeItems) {
                            $Results.IncompleteMoveReason += "Previous migration encountered $($moveHistory.LargeItemsEncountered) large items"
                            $Results.Warnings += "Previous migration encountered $($moveHistory.LargeItemsEncountered) large items"
                            Write-Log -Message "Warning: Previous migration for $EmailAddress encountered $($moveHistory.LargeItemsEncountered) large items" -Level "WARNING"
                        }
                        
                        Write-Log -Message "Recommendation: Consider using a higher BadItemLimit for this mailbox" -Level "INFO"
                    }
                }
            }
            catch [Microsoft.Exchange.Management.Migration.MoveRequestNotFoundException] {
                # No move history found, this is fine
            }
            catch {
                Write-Log -Message "Error checking migration history: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "WARNING"
                $Results.Warnings += "Error checking migration history: $($_.Exception.Message)"
            }
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Unexpected error in Test-MoveRequests: $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level "ERROR"
        if ($_.ScriptStackTrace) {
            Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level "DEBUG"
        }
        $Results.Warnings += "Failed to check move requests: $_"
        return $false
    }
}
