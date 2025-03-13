function Export-HTMLReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$TestResults,
        
        [Parameter(Mandatory = $false)]
        [string]$TemplatePath = ".\Templates\ReportTemplate.html",
        
        [Parameter(Mandatory = $false)]
        [string]$ReportPath,
        
        [Parameter(Mandatory = $false)]
        [string]$BatchName,
        
        [Parameter(Mandatory = $false)]
        [object]$MigrationBatch = $null
    )
    
    try {
        # Set default report path if not specified
        if (-not $ReportPath) {
            $ReportPath = Join-Path -Path $script:Config.ReportPath -ChildPath "$BatchName-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
        }
        
        # Check if template file exists
        if (-not (Test-Path -Path $TemplatePath)) {
            Write-Log -Message "HTML template not found: $TemplatePath" -Level "ERROR"
            Write-Log -Message "Using default inline template" -Level "WARNING"
            
            # Use embedded fallback template
            $templateContent = Get-EmbeddedHTMLTemplate
        }
        else {
            # Load template from file
            $templateContent = Get-Content -Path $TemplatePath -Raw -Encoding UTF8
            Write-Log -Message "Using HTML template: $TemplatePath" -Level "INFO"
        }
        
        # Replace template variables
        $reportContent = $templateContent -replace '{{BatchName}}', $BatchName
        $reportContent = $reportContent -replace '{{ReportDate}}', (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        $reportContent = $reportContent -replace '{{ScriptVersion}}', $script:ScriptVersion
        
        # Calculate summary data
        $readyCount = ($TestResults | Where-Object { $_.OverallStatus -eq "Ready" }).Count
        $warningCount = ($TestResults | Where-Object { $_.OverallStatus -eq "Warning" }).Count
        $failedCount = ($TestResults | Where-Object { $_.OverallStatus -eq "Failed" }).Count
        
        $reportContent = $reportContent -replace '{{ReadyCount}}', $readyCount
        $reportContent = $reportContent -replace '{{WarningCount}}', $warningCount
        $reportContent = $reportContent -replace '{{FailedCount}}', $failedCount
        $reportContent = $reportContent -replace '{{TotalCount}}', $TestResults.Count
        
        # Add migration batch details if available
        if ($MigrationBatch) {
            $batchStatus = if ($MigrationBatch.IsDryRun) { "Dry Run - Not Created" } else { $MigrationBatch.Status }
            
            $batchDetails = @"
<p><strong>Migration Batch Status:</strong> $batchStatus</p>
<p><strong>Target Delivery Domain:</strong> $($script:Config.TargetDeliveryDomain)</p>
<p><strong>Migration Endpoint:</strong> $($script:Config.MigrationEndpointName)</p>
"@
            $reportContent = $reportContent -replace '{{BatchDetails}}', $batchDetails
        }
        else {
            $reportContent = $reportContent -replace '{{BatchDetails}}', ''
        }
        
        # Build summary table rows
        $summaryRows = ""
        foreach ($result in $TestResults) {
            $statusClass = switch ($result.OverallStatus) {
                "Ready" { "status-ready" }
                "Warning" { "status-warning" }
                "Failed" { "status-failed" }
                default { "" }
            }
            
            $issueCount = $result.Errors.Count + $result.Warnings.Count
            $specialType = if ($result.IsSpecialMailbox -eq $true) { $result.SpecialMailboxType } else { "Standard" }
            $lastLogon = if ($result.LastLogonTime) { (Get-Date $result.LastLogonTime -Format "yyyy-MM-dd") } else { "Never" }
            
            $summaryRows += @"
<tr>
    <td>$($result.EmailAddress)</td>
    <td>$($result.DisplayName)</td>
    <td>$($result.MailboxSizeGB)</td>
    <td>$($result.TotalItemCount)</td>
    <td><span class="$statusClass">$($result.OverallStatus)</span></td>
    <td>$issueCount</td>
    <td>$specialType</td>
    <td>$lastLogon</td>
</tr>
"@
        }
        $reportContent = $reportContent -replace '{{SummaryTableRows}}', $summaryRows
        
        # Build detailed mailbox sections
        $detailedResults = ""
        foreach ($result in $TestResults) {
            $statusClass = switch ($result.OverallStatus) {
                "Ready" { "status-ready" }
                "Warning" { "status-warning" }
                "Failed" { "status-failed" }
                default { "" }
            }
            
            # Start mailbox details section
            $detailedResults += @"
<div class="mailbox-details">
    <h3>$($result.DisplayName) <small>($($result.EmailAddress))</small> - <span class="$statusClass">$($result.OverallStatus)</span></h3>
    <div class="mailbox-details-content">
"@
            
            # Add basic properties
            $detailedResults += Add-DetailRow -Label "UPN" -Value $result.UPN
            $detailedResults += Add-DetailRow -Label "Recipient Type" -Value $result.RecipientTypeDetails
            $detailedResults += Add-DetailRow -Label "Mailbox Size" -Value "$($result.MailboxSizeGB) GB"
            $detailedResults += Add-DetailRow -Label "UPN Matches Primary SMTP" -Value $result.UPNMatchesPrimarySMTP
            $detailedResults += Add-DetailRow -Label "Has OnMicrosoft Address" -Value $result.HasOnMicrosoftAddress
            $detailedResults += Add-DetailRow -Label "All Domains Verified" -Value $result.AllDomainsVerified
            $detailedResults += Add-DetailRow -Label "Has Exchange License" -Value $result.HasExchangeLicense
            $detailedResults += Add-DetailRow -Label "License Details" -Value $result.LicenseDetails
            $detailedResults += Add-DetailRow -Label "License Provisioning Status" -Value $result.LicenseProvisioningStatus
            $detailedResults += Add-DetailRow -Label "Litigation Hold" -Value $result.LitigationHoldEnabled
            $detailedResults += Add-DetailRow -Label "Retention Hold" -Value $result.RetentionHoldEnabled
            $detailedResults += Add-DetailRow -Label "Archive Status" -Value $result.ArchiveStatus
            $detailedResults += Add-DetailRow -Label "Forwarding Enabled" -Value $result.ForwardingEnabled
            $detailedResults += Add-DetailRow -Label "Exchange GUID" -Value $result.ExchangeGuid
            $detailedResults += Add-DetailRow -Label "Has Legacy Exchange DN" -Value $result.HasLegacyExchangeDN
            $detailedResults += Add-DetailRow -Label "Pending Move Request" -Value $result.PendingMoveRequest
            
            # Add move request status if applicable
            if ($result.PendingMoveRequest) {
                $detailedResults += Add-DetailRow -Label "Move Request Status" -Value $result.MoveRequestStatus
            }
            
            # Add permissions if any
            if ($result.HasSendAsPermissions -or $result.FullAccessDelegates.Count -gt 0 -or $result.SendOnBehalfDelegates.Count -gt 0) {
                $permissionsValue = ""
                
                if ($result.HasSendAsPermissions) {
                    $permissionsValue += "<strong>Send As:</strong> $($result.SendAsPermissions -join ", ")<br>"
                }
                
                if ($result.FullAccessDelegates.Count -gt 0) {
                    $permissionsValue += "<strong>Full Access:</strong> $($result.FullAccessDelegates -join ", ")<br>"
                }
                
                if ($result.SendOnBehalfDelegates.Count -gt 0) {
                    $permissionsValue += "<strong>Send On Behalf:</strong> $($result.SendOnBehalfDelegates -join ", ")"
                }
                
                $detailedResults += Add-DetailRow -Label "Permissions" -Value $permissionsValue
            }
            
            # Add special mailbox information
            $detailedResults += @"
<div class="category">
    <h4>Special Mailbox Information</h4>
    $( Add-DetailRow -Label "Is Special Mailbox" -Value $result.IsSpecialMailbox )
"@
            
            if ($result.IsSpecialMailbox) {
                $detailedResults += Add-DetailRow -Label "Special Mailbox Type" -Value $result.SpecialMailboxType
                $detailedResults += Add-DetailRow -Label "Migration Guidance" -Value $result.SpecialMailboxGuidance
            }
            
            $detailedResults += @"
</div>
<div class="category">
    <h4>Mailbox Activity Information</h4>
    $( Add-DetailRow -Label "Last Logon Time" -Value $result.LastLogonTime )
    $( Add-DetailRow -Label "Is Inactive" -Value $result.IsInactive )
"@
            
            if ($result.IsInactive) {
                $detailedResults += Add-DetailRow -Label "Inactive Days" -Value $result.InactiveDays
            }
            
            $detailedResults += @"
</div>
<div class="category">
    <h4>Data Integrity Information</h4>
    $( Add-DetailRow -Label "Corruption Risk" -Value $result.PotentialCorruptionRisk )
    $( Add-DetailRow -Label "Recommended BadItemLimit" -Value $result.RecommendedBadItemLimit )
    $( Add-DetailRow -Label "Has Incomplete Moves" -Value $result.HasIncompleteMoves )
"@
            
            if ($result.HasIncompleteMoves -and $result.IncompleteMovesDetails) {
                $moveDetailsValue = "<ul>"
                foreach ($moveDetail in $result.IncompleteMovesDetails) {
                    $moveDetailsValue += "<li>Completion Time: $($moveDetail.CompletionTime), Bad Items: $($moveDetail.BadItems), Large Items: $($moveDetail.LargeItems)</li>"
                }
                $moveDetailsValue += "</ul>"
                
                $detailedResults += Add-DetailRow -Label "Incomplete Moves Details" -Value $moveDetailsValue
            }
            
            $detailedResults += @"
</div>
<div class="category">
    <h4>Extended Validation Results</h4>
    $( Add-DetailRow -Label "Total Items" -Value $result.TotalItemCount )
    $( Add-DetailRow -Label "Folder Count" -Value $result.FolderCount )
    $( Add-DetailRow -Label "Unified Messaging Enabled" -Value $result.UMEnabled )
    $( Add-DetailRow -Label "Has Large Items (>150MB)" -Value $result.HasLargeItems )
    $( Add-DetailRow -Label "Has Deeply Nested Folders" -Value $result.HasDeepFolderHierarchy )
    $( Add-DetailRow -Label "Has Orphaned Permissions" -Value $result.HasOrphanedPermissions )
    $( Add-DetailRow -Label "Security Group Memberships" -Value "Direct: $($result.DirectGroupCount), Nested: $($result.NestedGroupCount)" )
    $( Add-DetailRow -Label "Calendar Items" -Value "$($result.CalendarItemCount) items in $($result.CalendarFolderCount) folders" )
    $( Add-DetailRow -Label "Has Shared Calendars" -Value $result.HasSharedCalendars )
    $( Add-DetailRow -Label "Contact Items" -Value "$($result.ContactItemCount) items in $($result.ContactFolderCount) folders" )
    $( Add-DetailRow -Label "Is Arbitration Mailbox" -Value $result.IsArbitrationMailbox )
    $( Add-DetailRow -Label "Is Audit Log Mailbox" -Value $result.IsAuditLogMailbox )
</div>
"@
            
            # Show large items if any exist
            if ($result.HasLargeItems -and $result.LargeItemsDetails.Count -gt 0) {
                $detailedResults += @"
<div class="category">
    <h4>Large Items Details</h4>
    <table>
        <thead>
            <tr>
                <th>Folder Path</th>
                <th>Max Item Size</th>
            </tr>
        </thead>
        <tbody>
"@
                foreach ($item in $result.LargeItemsDetails) {
                    $detailedResults += @"
            <tr>
                <td>$($item.FolderPath)</td>
                <td>$($item.MaxItemSize)</td>
            </tr>
"@
                }
                $detailedResults += @"
        </tbody>
    </table>
</div>
"@
            }
            
            # Add errors and warnings
            if ($result.Errors.Count -gt 0 -or $result.Warnings.Count -gt 0) {
                $issuesValue = ""
                
                if ($result.Errors.Count -gt 0) {
                    $issuesValue += "<strong>Errors:</strong><ul class='error-list'>"
                    
                    for ($i = 0; $i -lt $result.Errors.Count; $i++) {
                        $errorCode = if ($i -lt $result.ErrorCodes.Count) { $result.ErrorCodes[$i] } else { "" }
                        $errorCodeHtml = if ($errorCode) { "<span class='error-code'>$errorCode</span>" } else { "" }
                        
                        $issuesValue += "<li>$errorCodeHtml$($result.Errors[$i])</li>"
                    }
                    
                    $issuesValue += "</ul>"
                }
                
                if ($result.Warnings.Count -gt 0) {
                    $issuesValue += "<strong>Warnings:</strong><ul class='warning-list'>"
                    
                    foreach ($warning in $result.Warnings) {
                        $issuesValue += "<li>$warning</li>"
                    }
                    
                    $issuesValue += "</ul>"
                }
                
                $detailedResults += Add-DetailRow -Label "Issues" -Value $issuesValue
            }
            
            # Close mailbox details section
            $detailedResults += @"
    </div>
</div>
"@
        }
        $reportContent = $reportContent -replace '{{DetailedResults}}', $detailedResults
        
        # Build migration guidance
        $criticalIssues = @()
        foreach ($result in $TestResults) {
            if ($result.Errors.Count -gt 0) {
                $criticalIssues += "<li><strong>$($result.DisplayName) ($($result.EmailAddress))</strong>: $($result.Errors[0])</li>"
            }
        }
        
        if ($criticalIssues.Count -gt 0) {
            $criticalIssuesContent = $criticalIssues -join "`n"
        }
        else {
            $criticalIssuesContent = "<li>No critical issues found. All mailboxes can be migrated.</li>"
        }
        $reportContent = $reportContent -replace '{{CriticalIssues}}', $criticalIssuesContent
        
        # Build performance considerations
        $performanceIssues = @()
        foreach ($result in $TestResults) {
            if ($result.MailboxSizeGB -gt $script:Config.MaxMailboxSizeGB) {
                $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: Large mailbox size ($($result.MailboxSizeGB) GB)</li>"
            }
            if ($result.HasLargeFolders) {
                $folderCount = ($result.LargeFolders | Where-Object { $_.ItemsInFolder -gt 50000 }).Count
                if ($folderCount -gt 0) {
                    $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: Has $folderCount folders with over 50,000 items</li>"
                }
            }
            if ($result.TotalItemCount -gt 100000) {
                $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: High item count ($($result.TotalItemCount) items)</li>"
            }
            if ($result.HasLargeItems) {
                $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: Contains items larger than 150 MB</li>"
            }
            if ($result.PotentialCorruptionRisk -eq "High") {
                $performanceIssues += "<li><strong>$($result.DisplayName)</strong>: High risk of corrupted items (Recommended BadItemLimit: $($result.RecommendedBadItemLimit))</li>"
            }
        }
        
        if ($performanceIssues.Count -gt 0) {
            $performanceIssuesContent = $performanceIssues -join "`n"
        }
        else {
            $performanceIssuesContent = "<li>No performance issues found.</li>"
        }
        $reportContent = $reportContent -replace '{{PerformanceIssues}}', $performanceIssuesContent
        
        # Build special mailboxes
        $specialMailboxes = @()
        foreach ($result in $TestResults) {
            if ($result.IsSpecialMailbox) {
                $specialMailboxes += "<li><strong>$($result.DisplayName) ($($result.EmailAddress))</strong>: $($result.SpecialMailboxType) - $($result.SpecialMailboxGuidance)</li>"
            }
        }
        
        if ($specialMailboxes.Count -gt 0) {
            $specialMailboxesContent = $specialMailboxes -join "`n"
        }
        else {
            $specialMailboxesContent = "<li>No special mailbox types found.</li>"
        }
        $reportContent = $reportContent -replace '{{SpecialMailboxes}}', $specialMailboxesContent
        
        # Build post-migration tasks
        $postMigrationTasks = @()
        foreach ($result in $TestResults) {
            if ($result.UMEnabled) {
                $postMigrationTasks += "<li><strong>$($result.DisplayName)</strong>: Configure Cloud Voicemail to replace Unified Messaging</li>"
            }
            if ($result.HasSharedCalendars) {
                $postMigrationTasks += "<li><strong>$($result.DisplayName)</strong>: Recreate calendar sharing permissions</li>"
            }
            if ($result.NestedGroupCount -gt 0) {
                $postMigrationTasks += "<li><strong>$($result.DisplayName)</strong>: Verify nested group memberships</li>"
            }
            if ($result.HasIncompleteMoves) {
                $postMigrationTasks += "<li><strong>$($result.DisplayName)</strong>: Verify mailbox content completeness due to previous incomplete migration</li>"
            }
        }
        
        if ($postMigrationTasks.Count -gt 0) {
            $postMigrationTasksContent = $postMigrationTasks -join "`n"
        }
        else {
            $postMigrationTasksContent = "<li>No specific post-migration tasks identified.</li>"
        }
        $reportContent = $reportContent -replace '{{PostMigrationTasks}}', $postMigrationTasksContent
        
        # Add action required section if there are failed mailboxes
        $failedMailboxes = $TestResults | Where-Object { $_.OverallStatus -eq "Failed" }
        if ($failedMailboxes.Count -gt 0) {
            $actionRequiredContent = @"
<div class="action-required">
    <h3>Action Required</h3>
    <p>The following mailboxes have issues that need to be resolved before migration:</p>
    <ul>
"@
            
            foreach ($mailbox in $failedMailboxes) {
                $actionRequiredContent += @"
        <li><strong>$($mailbox.DisplayName) ($($mailbox.EmailAddress))</strong> - Issues: $($mailbox.Errors.Count)</li>
"@
            }
            
            $actionRequiredContent += @"
    </ul>
</div>
"@
            $reportContent = $reportContent -replace '{{ActionRequired}}', $actionRequiredContent
        }
        else {
            $reportContent = $reportContent -replace '{{ActionRequired}}', ''
        }
        
        # Save HTML to file
        $reportContent | Out-File -FilePath $ReportPath -Encoding utf8
        
        Write-Log -Message "HTML report generated: $ReportPath" -Level "SUCCESS"
        return $ReportPath
    }
    catch {
        Write-Log -Message "Failed to generate HTML report: $_" -Level "ERROR"
        return $null
    }
}

# Helper function to add detail rows to the report
function Add-DetailRow {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Label,
        
        [Parameter(Mandatory = $true)]
        $Value
    )
    
    return @"
<div class="detail-row">
    <div class="detail-label">$Label</div>
    <div class="detail-value">$Value</div>
</div>
"@
}

# Fallback HTML template in case the external template is not available
function Get-EmbeddedHTMLTemplate {
    [CmdletBinding()]
    param()
    
    return @'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Online Migration Report - {{BatchName}}</title>
    <style>
        /* Basic styles for fallback template */
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 20px; }
        .container { max-width: 1200px; margin: 0 auto; }
        h1, h2, h3 { color: #0078d4; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th, td { text-align: left; padding: 8px; border-bottom: 1px solid #ddd; }
        th { background-color: #0078d4; color: white; }
        .status-ready { background-color: #dff0d8; color: #3c763d; padding: 3px 8px; border-radius: 3px; }
        .status-warning { background-color: #fcf8e3; color: #8a6d3b; padding: 3px 8px; border-radius: 3px; }
        .status-failed { background-color: #f2dede; color: #a94442; padding: 3px 8px; border-radius: 3px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Exchange Online Migration Report</h1>
        
        <div>
            <h2>Migration Batch: {{BatchName}}</h2>
            <p><strong>Report Generated:</strong> {{ReportDate}}</p>
            <p><strong>Total Mailboxes:</strong> {{TotalCount}}</p>
            <p><strong>Ready for Migration:</strong> {{ReadyCount}}</p>
            <p><strong>Warnings:</strong> {{WarningCount}}</p>
            <p><strong>Failed:</strong> {{FailedCount}}</p>
            <p><strong>Script Version:</strong> {{ScriptVersion}}</p>
            {{BatchDetails}}
        </div>
        
        <h2>Mailbox Summary</h2>
        <table>
            <thead>
                <tr>
                    <th>Email Address</th>
                    <th>Display Name</th>
                    <th>Mailbox Size (GB)</th>
                    <th>Items</th>
                    <th>Status</th>
                    <th>Issues</th>
                    <th>Special Type</th>
                    <th>Last Logon</th>
                </tr>
            </thead>
            <tbody>
                {{SummaryTableRows}}
            </tbody>
        </table>
        
        <h2>Detailed Results</h2>
        {{DetailedResults}}
        
        <h2>Migration Guidance</h2>
        
        <h3>Critical Issues</h3>
        <ul>
            {{CriticalIssues}}
        </ul>
        
        <h3>Performance Considerations</h3>
        <ul>
            {{PerformanceIssues}}
        </ul>
        
        <h3>Special Mailbox Types</h3>
        <ul>
            {{SpecialMailboxes}}
        </ul>
        
        <h3>Post-Migration Tasks</h3>
        <ul>
            {{PostMigrationTasks}}
        </ul>
        
        {{ActionRequired}}
        
        <p style="margin-top: 30px; text-align: center; color: #777; font-style: italic;">
            Generated by Exchange Online Migration Script v{{ScriptVersion}}
        </p>
    </div>
</body>
</html>
'@
}
