function Test-OnPremCalendarAndContactItems {
    <#
    .SYNOPSIS
        Tests calendar and contact items in an on-premises mailbox for migration readiness.
    
    .DESCRIPTION
        Analyzes calendar and contact items in an on-premises mailbox to identify potential
        migration issues such as shared calendars, high meeting counts, or
        permissions that may need special handling during migration.
    
    .PARAMETER EmailAddress
        The email address of the mailbox to test.
    
    .PARAMETER Results
        A PSCustomObject that collects the validation results.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
        Test-OnPremCalendarAndContactItems -EmailAddress "user@contoso.com" -Results $results
    
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
        Write-Log -Message "Checking on-premises calendar and contact items: $EmailAddress" -Level "INFO"
        
        # Get folder statistics
        $folderStats = Get-MailboxFolderStatistics -Identity $EmailAddress
        
        # Check calendar folders - using FolderType property which is locale-independent
        $calendarFolders = $folderStats | Where-Object { $_.FolderType -eq "Calendar" }
        $Results.CalendarFolderCount = $calendarFolders.Count
        $Results.CalendarItemCount = ($calendarFolders | Measure-Object -Property ItemsInFolder -Sum).Sum
        
        # Check for shared calendars
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction Stop
        $calendarPermissions = @()
        
        # Get default calendar folder name (which may be localized)
        $defaultCalendarFolder = $calendarFolders | Where-Object { 
            # Check for default calendar paths in different languages
            $_.FolderPath -eq "/Calendar" -or 
            $_.FolderPath -eq "/Agenda" -or  # Dutch
            $_.FolderPath -eq "/Kalender" -or # German
            $_.FolderPath -eq "/Calendrier" -or # French
            $_.IsDefaultFolder -eq $true
        }
        
        # Log detected default calendar folder for debugging
        if ($defaultCalendarFolder) {
            Write-Log -Message "Detected default calendar folder: $($defaultCalendarFolder.FolderPath)" -Level "DEBUG"
        }
        
        foreach ($calFolder in $calendarFolders) {
            # Use the path with escaped backslashes for PowerShell
            $folderPath = $calFolder.FolderPath.Replace('/', '\')
            
            # Handle different ways to construct folder identity
            try {
                # First try direct folder identity construction
                $folderIdentity = "$($mailbox.PrimarySmtpAddress):$folderPath"
                $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop
            }
            catch {
                try {
                    # If the first approach fails, try with the folder ID
                    $folderIdentity = "$($mailbox.PrimarySmtpAddress):\$($calFolder.FolderId)"
                    $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop
                }
                catch {
                    try {
                        # If that fails, try with the default format but removing the leading backslash
                        $folderIdentity = "$($mailbox.PrimarySmtpAddress):$($folderPath.TrimStart('\'))"
                        $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop
                    }
                    catch {
                        Write-Log -Message "Could not check permissions for calendar folder $($calFolder.FolderPath): $_" -Level "DEBUG"
                        continue
                    }
                }
            }
            
            # Process permissions if we got them
            $nonDefaultPermissions = $permissions | Where-Object { 
                $_.User.DisplayName -ne "Default" -and 
                $_.User.DisplayName -ne "Anonymous" -and
                $_.AccessRights -ne "None"
            }
            
            if ($nonDefaultPermissions) {
                foreach ($perm in $nonDefaultPermissions) {
                    $calendarPermissions += [PSCustomObject]@{
                        FolderPath = $calFolder.FolderPath
                        User = $perm.User.DisplayName
                        AccessRights = $perm.AccessRights -join ", "
                    }
                }
            }
        }
        
        if ($calendarPermissions.Count -gt 0) {
            $Results.HasSharedCalendars = $true
            $Results.CalendarPermissions = $calendarPermissions
            $Results.Warnings += "Mailbox has shared calendars with custom permissions that need to be recreated post-migration"
            $Results.ErrorCodes += "ERR025"
            
            Write-Log -Message "Warning: Mailbox $EmailAddress has shared calendars with custom permissions:" -Level "WARNING" -ErrorCode "ERR025"
            foreach ($perm in ($calendarPermissions | Select-Object -First 5)) {
                Write-Log -Message "  - $($perm.FolderPath): Shared with $($perm.User) ($($perm.AccessRights))" -Level "WARNING"
            }
            
            if ($calendarPermissions.Count -gt 5) {
                Write-Log -Message "  - ... and $($calendarPermissions.Count - 5) more calendar sharing permissions" -Level "WARNING"
            }
            
            Write-Log -Message "Recommendation: Document calendar sharing permissions as they may need to be manually recreated" -Level "INFO"
        }
        else {
            $Results.HasSharedCalendars = $false
            Write-Log -Message "No shared calendars found for mailbox $EmailAddress" -Level "INFO"
        }
        
        # Check for large number of recurring meetings
        if ($Results.CalendarItemCount -gt 1000) {
            $Results.Warnings += "Mailbox has a large number of calendar items ($($Results.CalendarItemCount)), which may include many recurring meetings"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a large number of calendar items: $($Results.CalendarItemCount)" -Level "WARNING"
            Write-Log -Message "Recommendation: Consider cleaning up old calendar items before migration" -Level "INFO"
        }
        
        # Check for calendar processing settings (for room mailboxes)
        if ($mailbox.RecipientTypeDetails -eq "RoomMailbox" -or $mailbox.RecipientTypeDetails -eq "EquipmentMailbox") {
            try {
                $calendarProcessing = Get-CalendarProcessing -Identity $mailbox.Identity
                
                if ($calendarProcessing) {
                    # Check for custom calendar processing settings
                    $hasCustomSettings = 
                        $calendarProcessing.AutomateProcessing -ne "AutoAccept" -or
                        $calendarProcessing.AllowConflicts -eq $true -or
                        $calendarProcessing.BookingWindowInDays -ne 180 -or
                        $calendarProcessing.MaximumDurationInMinutes -ne 1440
                        
                    if ($hasCustomSettings) {
                        $Results.Warnings += "Room mailbox has custom calendar processing settings that need to be recreated post-migration"
                        Write-Log -Message "Warning: Room mailbox $EmailAddress has custom calendar processing settings" -Level "WARNING"
                        Write-Log -Message "Recommendation: Document these settings and reconfigure them in Exchange Online after migration" -Level "INFO"
                    }
                    
                    # Check for calendar delegates
                    if ($calendarProcessing.ResourceDelegates -and $calendarProcessing.ResourceDelegates.Count -gt 0) {
                        $Results.Warnings += "Room mailbox has custom resource delegates that need to be recreated post-migration"
                        Write-Log -Message "Warning: Room mailbox $EmailAddress has the following resource delegates:" -Level "WARNING"
                        foreach ($delegate in $calendarProcessing.ResourceDelegates) {
                            Write-Log -Message "  - $delegate" -Level "WARNING"
                        }
                        Write-Log -Message "Recommendation: Document these delegates and reconfigure them in Exchange Online after migration" -Level "INFO"
                    }
                }
            }
            catch {
                Write-Log -Message "Failed to get calendar processing information: $_" -Level "WARNING"
            }
        }
        
        # Check contacts folders
        $contactFolders = $folderStats | Where-Object { $_.FolderType -eq "Contacts" }
        $Results.ContactFolderCount = $contactFolders.Count
        $Results.ContactItemCount = ($contactFolders | Measure-Object -Property ItemsInFolder -Sum).Sum
        
        if ($Results.ContactItemCount -gt 5000) {
            $Results.Warnings += "Mailbox has a large number of contacts ($($Results.ContactItemCount))"
            Write-Log -Message "Warning: Mailbox $EmailAddress has a large number of contacts: $($Results.ContactItemCount)" -Level "WARNING"
            Write-Log -Message "Recommendation: Consider exporting contacts to a CSV before migration" -Level "INFO"
        }
        
        return $true
    }
    catch {
        $Results.Warnings += "Failed to check calendar and contact items: $_"
        Write-Log -Message "Warning: Failed to check calendar and contact items for $EmailAddress`: $_" -Level "WARNING"
        return $false
    }
}