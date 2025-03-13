function Add-DetailRow {
    <#
    .SYNOPSIS
        Creates an HTML detail row for the migration report.
    
    .DESCRIPTION
        Generates HTML markup for a label-value pair in the migration report.
        Used to build detail sections for each mailbox in the HTML report.
    
    .PARAMETER Label
        The label for the detail row.
    
    .PARAMETER Value
        The value to display for the detail row.
    
    .PARAMETER HighlightPositive
        When specified, applies positive highlighting to boolean true values.
    
    .PARAMETER HighlightNegative
        When specified, applies negative highlighting to boolean false values.
    
    .EXAMPLE
        Add-DetailRow -Label "UPN" -Value $Results.UPN
    
    .EXAMPLE
        Add-DetailRow -Label "Has Exchange License" -Value $Results.HasExchangeLicense -HighlightPositive -HighlightNegative
    
    .OUTPUTS
        [string] Returns HTML markup for the detail row.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Label,
        
        [Parameter(Mandatory = $true)]
        $Value,
        
        [Parameter(Mandatory = $false)]
        [switch]$HighlightPositive,
        
        [Parameter(Mandatory = $false)]
        [switch]$HighlightNegative
    )
    
    # Format the value based on its type
    $formattedValue = $Value
    
    if ($Value -is [System.Boolean] -or $Value -is [bool]) {
        # Apply highlighting for boolean values if requested
        if ($Value -and $HighlightPositive) {
            $formattedValue = "<span class='yes-value'>$Value</span>"
        }
        elseif (-not $Value -and $HighlightNegative) {
            $formattedValue = "<span class='no-value'>$Value</span>"
        }
    }
    elseif ($Value -is [System.DateTime]) {
        # Format date/time values
        $formattedValue = $Value.ToString("yyyy-MM-dd HH:mm:ss")
    }
    elseif ($Value -is [System.Array] -or $Value -is [System.Collections.ArrayList]) {
        # Format arrays
        if ($Value.Count -eq 0) {
            $formattedValue = "None"
        }
        else {
            $formattedValue = $Value -join ", "
        }
    }
    elseif ($null -eq $Value) {
        # Handle null values
        $formattedValue = "None"
    }
    
    # Generate the HTML for the detail row
    return @"
<div class="detail-row">
    <div class="detail-label">$Label</div>
    <div class="detail-value">$formattedValue</div>
</div>
"@
}
