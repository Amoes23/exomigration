function Send-MigrationNotification {
    <#
    .SYNOPSIS
        Sends email notifications about migration status and events.
    
    .DESCRIPTION
        Sends formatted email notifications to administrators or stakeholders
        regarding migration status, events, or errors. This function supports
        both plain text and HTML email formats and can use stored credentials.
    
    .PARAMETER Subject
        The subject line for the email notification.
    
    .PARAMETER Body
        The body content of the email notification.
    
    .PARAMETER To
        Email addresses to send the notification to. Defaults to the NotificationEmails
        from the configuration if not specified.
    
    .PARAMETER From
        The sender's email address. Defaults to a generated address if not specified.
    
    .PARAMETER SmtpServer
        The SMTP server to use for sending. Defaults to smtp.office365.com if not specified.
    
    .PARAMETER Port
        The SMTP port to use. Defaults to 587 (TLS) if not specified.
    
    .PARAMETER Credential
        Optional PSCredential object containing SMTP authentication credentials.
        If not provided, will prompt if needed.
    
    .PARAMETER UseSsl
        When specified, uses SSL/TLS for the SMTP connection. Default is true.
    
    .PARAMETER BodyAsHtml
        When specified, formats the email body as HTML. Default is true.
    
    .PARAMETER Priority
        Email priority: Normal, High, or Low. Default is Normal.
    
    .PARAMETER Attachments
        Optional array of file paths to attach to the email.
    
    .EXAMPLE
        Send-MigrationNotification -Subject "Migration Started" -Body "Migration batch started successfully."
    
    .EXAMPLE
        Send-MigrationNotification -Subject "Migration Error" -Body "<h1>Migration Failed</h1><p>Error details: $errorMsg</p>" -Priority High
    
    .OUTPUTS
        [bool] Returns $true if the email was sent successfully, $false otherwise.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        
        [Parameter(Mandatory = $true)]
        [string]$Body,
        
        [Parameter(Mandatory = $false)]
        [string[]]$To,
        
        [Parameter(Mandatory = $false)]
        [string]$From,
        
        [Parameter(Mandatory = $false)]
        [string]$SmtpServer = "smtp.office365.com",
        
        [Parameter(Mandatory = $false)]
        [int]$Port = 587,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $false)]
        [switch]$UseSsl = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$BodyAsHtml = $true,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Normal', 'High', 'Low')]
        [string]$Priority = 'Normal',
        
        [Parameter(Mandatory = $false)]
        [string[]]$Attachments
    )
    
    try {
        # Use config values if parameters not explicitly provided
        if (-not $To -or $To.Count -eq 0) {
            if ($script:Config -and $script:Config.NotificationEmails -and $script:Config.NotificationEmails.Count -gt 0) {
                $To = $script:Config.NotificationEmails
            }
            else {
                Write-Log -Message "No notification email recipients specified in parameters or configuration" -Level "WARNING"
                return $false
            }
        }
        
        if ([string]::IsNullOrEmpty($From)) {
            $computerName = $env:COMPUTERNAME
            $domain = $To[0].Split('@')[1]  # Use domain from first recipient
            $From = "ExchangeMigration@$computerName.$domain"
        }
        
        Write-Log -Message "Sending email notification: $Subject" -Level "INFO"
        Write-Verbose "To: $($To -join ', ')"
        
        # Set up email parameters
        $mailParams = @{
            To = $To
            From = $From
            Subject = $Subject
            Body = $Body
            SmtpServer = $SmtpServer
            Port = $Port
            UseSsl = $UseSsl
            BodyAsHtml = $BodyAsHtml
            Priority = $Priority
        }
        
        # Add attachments if specified
        if ($Attachments -and $Attachments.Count -gt 0) {
            $validAttachments = @()
            foreach ($attachment in $Attachments) {
                if (Test-Path -Path $attachment -PathType Leaf) {
                    $validAttachments += $attachment
                }
                else {
                    Write-Log -Message "Attachment not found: $attachment" -Level "WARNING"
                }
            }
            
            if ($validAttachments.Count -gt 0) {
                $mailParams.Add('Attachments', $validAttachments)
            }
        }
        
        # Handle credentials
        $useCredentials = $true
        if (-not $Credential) {
            # Check for stored credential
            if ($script:EmailCredential) {
                $Credential = $script:EmailCredential
            }
            else {
                # Prompt for credentials if required and save for future use
                try {
                    $Credential = Get-Credential -Message "Enter credentials for sending email notifications" -ErrorAction SilentlyContinue
                    if ($Credential) {
                        $script:EmailCredential = $Credential
                    }
                    else {
                        $useCredentials = $false
                        Write-Log -Message "No credentials provided for email, attempting without authentication" -Level "WARNING"
                    }
                }
                catch {
                    $useCredentials = $false
                    Write-Log -Message "Unable to prompt for credentials: $_" -Level "WARNING"
                }
            }
        }
        
        if ($useCredentials -and $Credential) {
            $mailParams.Add('Credential', $Credential)
        }
        
        # Send the email
        Send-MailMessage @mailParams
        
        Write-Log -Message "Email notification sent successfully to $($To -join ', ')" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log -Message "Failed to send email notification: $_" -Level "ERROR"
        Write-Verbose "Email parameters: Subject='$Subject', To='$($To -join ',')'"
        return $false
    }
}
