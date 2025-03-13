function New-MailboxTestResult {
    <#
    .SYNOPSIS
        Creates a mailbox test result object with default properties.
    
    .DESCRIPTION
        Creates a PSCustomObject with all properties required for validating a mailbox for migration.
        This ensures all validation functions have a consistent data structure to work with.
    
    .PARAMETER EmailAddress
        The email address of the mailbox being tested.
    
    .EXAMPLE
        $results = New-MailboxTestResult -EmailAddress "user@contoso.com"
    
    .OUTPUTS
        [PSCustomObject] Returns a PSCustomObject with default property values.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmailAddress
    )
    
    # Define all properties with default values
    $mailboxResultProperties = @(
        # Basic mailbox information
        @{Name = "EmailAddress"; DefaultValue = $null},
        @{Name = "DisplayName"; DefaultValue = $null},
        @{Name = "UPN"; DefaultValue = $null},
        @{Name = "ExistingMailbox"; DefaultValue = $false},
        @{Name = "RecipientType"; DefaultValue = $null},
        @{Name = "RecipientTypeDetails"; DefaultValue = $null},
        @{Name = "MailboxEnabled"; DefaultValue = $false},
        
        # Domain and addressing properties
        @{Name = "UPNMatchesPrimarySMTP"; DefaultValue = $false},
        @{Name = "HasOnMicrosoftAddress"; DefaultValue = $false},
        @{Name = "HasRequiredOnMicrosoftAddress"; DefaultValue = $false},
        @{Name = "AllDomainsVerified"; DefaultValue = $false},
        @{Name = "HasUnverifiedDomains"; DefaultValue = $false},
        @{Name = "UnverifiedDomains"; DefaultValue = @()},
        @{Name = "ProxyAddressesFormatting"; DefaultValue = $true},
        
        # License information
        @{Name = "HasExchangeLicense"; DefaultValue = $false},
        @{Name = "HasE1OrE5License"; DefaultValue = $false},
        @{Name = "LicenseType"; DefaultValue = $null},
        @{Name = "LicenseProvisioningStatus"; DefaultValue = $null},
        @{Name = "LicenseDetails"; DefaultValue = $null},
        
        # Mailbox configuration
        @{Name = "HasSendAsPermissions"; DefaultValue = $false},
        @{Name = "SendAsPermissions"; DefaultValue = @()},
        @{Name = "HiddenFromAddressList"; DefaultValue = $false},
        @{Name = "SAMAccountNameValid"; DefaultValue = $false},
        @{Name = "SAMAccountName"; DefaultValue = $null},
        @{Name = "AliasValid"; DefaultValue = $false},
        @{Name = "Alias"; DefaultValue = $null},
        @{Name = "LitigationHoldEnabled"; DefaultValue = $false},
        @{Name = "RetentionHoldEnabled"; DefaultValue = $false},
        @{Name = "MailboxSizeGB"; DefaultValue = 0},
        @{Name = "MailboxSizeWarning"; DefaultValue = $false},
        @{Name = "ArchiveStatus"; DefaultValue = $null},
        @{Name = "ForwardingEnabled"; DefaultValue = $false},
        @{Name = "ForwardingAddress"; DefaultValue = $null},
        @{Name = "ExchangeGuid"; DefaultValue = $null},
        @{Name = "HasLegacyExchangeDN"; DefaultValue = $false},
        
        # Delegation and permissions
        @{Name = "FullAccessDelegates"; DefaultValue = @()},
        @{Name = "SendOnBehalfDelegates"; DefaultValue = @()},
        
        # Move request information
        @{Name = "PendingMoveRequest"; DefaultValue = $false},
        @{Name = "MoveRequestStatus"; DefaultValue = $null},
        
        # Extended validation properties
        @{Name = "UMEnabled"; DefaultValue = $false},
        @{Name = "UMDetails"; DefaultValue = $null},
        @{Name = "HasLargeItems"; DefaultValue = $false},
        @{Name = "LargeItemsDetails"; DefaultValue = @()},
        @{Name = "HasOrphanedPermissions"; DefaultValue = $false},
        @{Name = "OrphanedPermissions"; DefaultValue = @()},
        @{Name = "GroupMemberships"; DefaultValue = @()},
        @{Name = "DirectGroupCount"; DefaultValue = 0},
        @{Name = "NestedGroupCount"; DefaultValue = 0},
        @{Name = "FolderCount"; DefaultValue = 0},
        @{Name = "HasDeepFolderHierarchy"; DefaultValue = $false},
        @{Name = "DeepFolders"; DefaultValue = @()},
        @{Name = "HasLargeFolders"; DefaultValue = $false},
        @{Name = "LargeFolders"; DefaultValue = @()},
        @{Name = "TotalItemCount"; DefaultValue = 0},
        @{Name = "CalendarFolderCount"; DefaultValue = 0},
        @{Name = "CalendarItemCount"; DefaultValue = 0},
        @{Name = "HasSharedCalendars"; DefaultValue = $false},
        @{Name = "CalendarPermissions"; DefaultValue = @()},
        @{Name = "ContactFolderCount"; DefaultValue = 0},
        @{Name = "ContactItemCount"; DefaultValue = 0},
        @{Name = "IsArbitrationMailbox"; DefaultValue = $false},
        @{Name = "IsAuditLogMailbox"; DefaultValue = $false},
        
        # Added in v3.0
        @{Name = "LastLogonTime"; DefaultValue = $null},
        @{Name = "IsInactive"; DefaultValue = $false},
        @{Name = "InactiveDays"; DefaultValue = 0},
        @{Name = "PotentialCorruptionRisk"; DefaultValue = "Low"},
        @{Name = "RecommendedBadItemLimit"; DefaultValue = 10},
        @{Name = "IsSpecialMailbox"; DefaultValue = $false},
        @{Name = "SpecialMailboxType"; DefaultValue = $null},
        @{Name = "SpecialMailboxGuidance"; DefaultValue = $null},
        @{Name = "HasIncompleteMoves"; DefaultValue = $false},
        @{Name = "IncompleteMovesDetails"; DefaultValue = $null},
        
        # New properties for additional validation
        @{Name = "LicenseSpecificQuotaGB"; DefaultValue = 50},
        @{Name = "MailboxNearQuota"; DefaultValue = $false},
        @{Name = "ADSyncVerified"; DefaultValue = $false},
        @{Name = "IsDirSynced"; DefaultValue = $false},
        @{Name = "LastDirSyncTime"; DefaultValue = $null},
        @{Name = "SyncIssues"; DefaultValue = @()},
        @{Name = "HasArchive"; DefaultValue = $false},
        @{Name = "ArchiveSizeGB"; DefaultValue = 0},
        @{Name = "ArchiveItemCount"; DefaultValue = 0},
        @{Name = "ArchiveAutoExpandingEnabled"; DefaultValue = $false},
        @{Name = "HasCloudPlaceholder"; DefaultValue = $false},
        @{Name = "CloudPlaceholderDetails"; DefaultValue = $null},
        @{Name = "AuditEnabled"; DefaultValue = $false},
        @{Name = "AuditLogAgeLimit"; DefaultValue = 90},
        @{Name = "AuditOwner"; DefaultValue = @()},
        @{Name = "AuditAdmin"; DefaultValue = @()},
        @{Name = "AuditDelegate"; DefaultValue = @()},
        @{Name = "CustomAuditSettings"; DefaultValue = $false},
        @{Name = "RecoverableItemsFolderSizeGB"; DefaultValue = 0},
        @{Name = "RecoverableItemsCount"; DefaultValue = 0},
        @{Name = "RecoverableItemsFolderDetails"; DefaultValue = @()},
        @{Name = "IsSoftDeleted"; DefaultValue = $false},
        @{Name = "HasSoftDeletedConflicts"; DefaultValue = $false},
        @{Name = "SoftDeletedConflicts"; DefaultValue = @()},
        @{Name = "HasAliasConflicts"; DefaultValue = $false},
        @{Name = "AliasConflicts"; DefaultValue = @()},
        @{Name = "IsJournalRecipient"; DefaultValue = $false},
        @{Name = "IsPartOfJournalRule"; DefaultValue = $false},
        @{Name = "JournalRules"; DefaultValue = @()},
        @{Name = "HasNamespaceConflict"; DefaultValue = $false},
        @{Name = "NamespaceConflictDetails"; DefaultValue = $null},
        @{Name = "HasFolderNameConflicts"; DefaultValue = $false},
        @{Name = "FolderNameConflicts"; DefaultValue = @()},
        @{Name = "HasOldItems"; DefaultValue = $false},
        @{Name = "OldestItemAge"; DefaultValue = 0},
        @{Name = "ItemAgeDistribution"; DefaultValue = $null},
        @{Name = "ActivityLevel"; DefaultValue = "Low"},
        @{Name = "AverageDailyEmailCount"; DefaultValue = 0},
        @{Name = "LastActive"; DefaultValue = $null},
        @{Name = "HasDeepFolderPermissions"; DefaultValue = $false},
        @{Name = "FolderPermissionsCount"; DefaultValue = 0},
        @{Name = "MaxFolderPermissionDepth"; DefaultValue = 0},
        @{Name = "ComplexPermissionFolders"; DefaultValue = @()},
        
        # Result tracking
        @{Name = "Warnings"; DefaultValue = @()},
        @{Name = "Errors"; DefaultValue = @()},
        @{Name = "ErrorCodes"; DefaultValue = @()},
        @{Name = "OverallStatus"; DefaultValue = "Unknown"},
        @{Name = "ValidationLevel"; DefaultValue = "Standard"}
    )
    
    # Create a result object with all properties initialized
    $result = [PSCustomObject]@{}
    
    # Add each property with its default value
    foreach ($prop in $mailboxResultProperties) {
        # If it's the EmailAddress property, use the provided value
        if ($prop.Name -eq "EmailAddress") {
            $result | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $EmailAddress
        }
        else {
            $result | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.DefaultValue
        }
    }
    
    return $result
}
