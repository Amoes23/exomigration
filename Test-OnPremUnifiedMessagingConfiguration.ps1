function Test-OnPremUnifiedMessagingConfiguration {
    <#
    .SYNOPSIS
        Tests if an on-premises mailbox has Unified Messaging features enabled.
    
    .DESCRIPTION
        Checks if an on-premises mailbox has Unified Messaging (UM) enabled, which requires
        special handling during migration. Identifies UM policies and settings
        that may need reconfiguration after migration to Exchange Online.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-OnPremUnifiedMessagingConfiguration -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking on-premises Unified Messaging configuration for: $EmailAddress" -Level "INFO"
        
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
                    PIN = "****" # Do not store actual PIN
                    PINExpired = $umMailbox.PINExpired
                    PINLastChanged = $umMailbox.PINLastChanged
                }
                
                $Results.Warnings += "Unified Messaging is enabled and requires special handling during migration"
                $Results.ErrorCodes += "ERR020"
                
                Write-Log -Message "Warning: Unified Messaging is enabled for $EmailAddress" -Level "WARNING" -ErrorCode "ERR020"
                Write-Log -Message "  UM Policy: $($umMailbox.UMMailboxPolicy)" -Level "WARNING"
                Write-Log -Message "  SIP Resource: $($umMailbox.SIPResourceIdentifier)" -Level "WARNING"
                Write-Log -Message "  Extensions: $($umMailbox.Extensions -join ', ')" -Level "WARNING"
                
                Write-Log -Message "Recommendation: Document UM settings and plan to enable Cloud Voicemail post-migration" -Level "INFO"
                Write-Log -Message "Recommendation: Disable UM before migration or use the -MoveVoicemailToTarget switch in New-MoveRequest" -Level "INFO"
                Write-Log -Message "Note: Unified Messaging in Exchange on-premises is being replaced by Cloud Voicemail in Exchange Online" -Level "INFO"
                
                # Check if there are UM auto attendant settings
                try {
                    $umConfig = Get-UMMailboxConfiguration -Identity $EmailAddress -ErrorAction SilentlyContinue
                    if ($umConfig) {
                        $Results.UMDetails | Add-Member -NotePropertyName "AutoAnswerEnabled" -NotePropertyValue $umConfig.AutoAnswerEnabled -Force
                        $Results.UMDetails | Add-Member -NotePropertyName "CallAnsweringRules" -NotePropertyValue ($umConfig.CallAnsweringRules.Count) -Force
                        
                        if ($umConfig.CallAnsweringRules -and $umConfig.CallAnsweringRules.Count -gt 0) {
                            $Results.Warnings += "Mailbox has custom UM call answering rules that will need to be recreated in Cloud Voicemail"
                            Write-Log -Message "Warning: Mailbox has $($umConfig.CallAnsweringRules.Count) custom UM call answering rules" -Level "WARNING"
                            Write-Log -Message "Recommendation: Document these rules before migration as they won't migrate automatically" -Level "INFO"
                        }
                        
                        if ($umConfig.OofGreetingEnabled -or $umConfig.PlayOnPhoneEnabled) {
                            $Results.Warnings += "Mailbox has custom UM greeting settings that need to be reconfigured in Cloud Voicemail"
                            Write-Log -Message "Warning: Mailbox has custom UM greeting settings" -Level "WARNING"
                        }
                    }
                }
                catch {
                    Write-Log -Message "Could not retrieve UM mailbox configuration: $_" -Level "WARNING"
                }
                
                # Check if there are existing voicemail messages
                try {
                    $voicemailFolder = $EmailAddress + ":\Voice Mail"
                    $voicemailStats = Get-MailboxFolderStatistics -Identity $voicemailFolder -ErrorAction SilentlyContinue
                    
                    if ($voicemailStats -and $voicemailStats.ItemsInFolder -gt 0) {
                        $Results.Warnings += "Mailbox has $($voicemailStats.ItemsInFolder) voicemail messages that need to be migrated"
                        Write-Log -Message "Warning: Mailbox has $($voicemailStats.ItemsInFolder) voicemail messages" -Level "WARNING"
                        Write-Log -Message "Recommendation: Use the -MoveVoicemailToTarget switch in New-MoveRequest to migrate voicemail messages" -Level "INFO"
                    }
                }
                catch {
                    Write-Log -Message "Could not check voicemail folder: $_" -Level "WARNING"
                }
            }
        }
        catch [System.Management.Automation.CommandNotFoundException] {
            # UM cmdlets not available in this environment
            Write-Log -Message "Unified Messaging cmdlets not available - checking recipient type" -Level "DEBUG"
            
            # Explicitly set to null to prevent access to non-existent object
            $umMailbox = $null
            
            # Check recipient type as a fallback
            try {
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
                Write-Log -Message "Error checking recipient properties: $_" -Level "WARNING"
                $Results.UMEnabled = $false
                $Results.Warnings += "Could not determine if Unified Messaging is enabled: $_"
            }
        }
        catch {
            # Handle other errors
            Write-Log -Message "Error checking Unified Messaging status: $_" -Level "WARNING"
            
            # Try checking mailbox directly
            try {
                $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
                if ($mailbox -and $mailbox.UMEnabled -eq $true) {
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
            catch {
                $Results.UMEnabled = $false
                $Results.Warnings += "Could not determine if Unified Messaging is enabled: $_"
                Write-Log -Message "Failed to check if Unified Messaging is enabled: $_" -Level "WARNING"
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