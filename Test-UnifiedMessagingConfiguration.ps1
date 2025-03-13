function Test-UnifiedMessagingConfiguration {
    <#
    .SYNOPSIS
        Tests if a mailbox has Unified Messaging features enabled.
    
    .DESCRIPTION
        Checks if a mailbox has Unified Messaging (UM) enabled, which requires
        special handling during migration. Identifies UM policies and settings
        that may need reconfiguration after migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-UnifiedMessagingConfiguration -EmailAddress "user@contoso.com" -Results $results
    
    .OUTPUTS
        [bool] Returns $true if the test was successful (even if issues were found), $false if the test failed.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Results
    )
    
    try {
        Write-Log -Message "Checking Unified Messaging configuration for: $EmailAddress" -Level "INFO"
        
        # Try to use Get-UMMailbox (if available in the environment)
        try {
            $umMailbox = Get-UMMailbox -Identity $EmailAddress -ErrorAction Stop
            if ($umMailbox) {
                $Results.UMEnabled = $true
                $Results.UMDetails = [PSCustomObject]@{
                    UMEnabled = $true
                    UMMailboxPolicy = $umMailbox.UMMailboxPolicy
                    SIPResourceIdentifier = $umMailbox.SIPResourceIdentifier
                    Extensions = $umMailbox.Extensions -join ", "
                    UMGrammar = $umMailbox.UMGrammar
                    UMRecipientDialPlanId = $umMailbox.UMRecipientDialPlanId
                    UMEnabledDate = $umMailbox.UMEnabledDate
                }
                
                $Results.Warnings += "Unified Messaging is enabled and requires special handling during migration"
                $Results.ErrorCodes += "ERR020"
                
                Write-Log -Message "Warning: Unified Messaging is enabled for $EmailAddress" -Level "WARNING" -ErrorCode "ERR020"
                Write-Log -Message "  UM Policy: $($umMailbox.UMMailboxPolicy)" -Level "WARNING"
                Write-Log -Message "  SIP Resource: $($umMailbox.SIPResourceIdentifier)" -Level "WARNING"
                Write-Log -Message "  Extensions: $($umMailbox.Extensions -join ', ')" -Level "WARNING"
                
                Write-Log -Message "Recommendation: Disable UM before migration and plan to enable Cloud Voicemail post-migration" -Level "INFO"
                Write-Log -Message "Documentation: https://docs.microsoft.com/en-us/exchange/voice-mail-unified-messaging/migrate-to-cloud-voicemail" -Level "INFO"
            }
        }
        catch [System.Management.Automation.CommandNotFoundException] {
            # UM cmdlets not available in this environment
            Write-Log -Message "Unified Messaging cmdlets not available - checking recipient type" -Level "DEBUG"
            
            # Explicitly set to null to prevent access to non-existent object
            $umMailbox = $null
            
            # Check recipient type as a fallback
            $recipient = Get-Recipient -Identity $EmailAddress -ErrorAction SilentlyContinue
            
            if ($recipient -and $recipient.RecipientTypeDetails -eq "UMEnabled") {
                $Results.UMEnabled = $true
                $Results.UMDetails = [PSCustomObject]@{
                    UMEnabled = $true
                    UMMailboxPolicy = "Unknown - use Exchange Admin Center to check"
                    Notes = "Detected via recipient type, use Exchange Admin Center for details"
                }
                
                $Results.Warnings += "Unified Messaging is enabled based on recipient type"
                $Results.ErrorCodes += "ERR020"
                
                Write-Log -Message "Warning: Unified Messaging appears to be enabled for $EmailAddress (detected via recipient type)" -Level "WARNING" -ErrorCode "ERR020"
                Write-Log -Message "Recommendation: Check Exchange Admin Center for UM details, disable UM before migration" -Level "INFO"
            }
            else {
                $Results.UMEnabled = $false
                Write-Log -Message "Unified Messaging is not enabled for $EmailAddress" -Level "INFO"
            }
        }
        catch {
            # Handle other errors
            Write-Log -Message "Error checking Unified Messaging status: $_" -Level "WARNING"
            
            # Check recipient type as a fallback
            $recipient = Get-Recipient -Identity $EmailAddress -ErrorAction SilentlyContinue
            
            if ($recipient -and $recipient.RecipientTypeDetails -eq "UMEnabled") {
                $Results.UMEnabled = $true
                $Results.UMDetails = [PSCustomObject]@{
                    UMEnabled = $true
                    UMMailboxPolicy = "Unknown"
                    Notes = "Error checking UM details: $_"
                }
                
                $Results.Warnings += "Unified Messaging appears to be enabled but details could not be retrieved"
                $Results.ErrorCodes += "ERR020"
                
                Write-Log -Message "Warning: Unified Messaging appears to be enabled for $EmailAddress but details could not be retrieved" -Level "WARNING" -ErrorCode "ERR020"
            }
            else {
                $Results.UMEnabled = $false
                Write-Log -Message "Unified Messaging is likely not enabled for $EmailAddress" -Level "INFO"
            }
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check Unified Messaging configuration: $_"
        Write-Log -Message "Warning: Failed to check Unified Messaging for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}
