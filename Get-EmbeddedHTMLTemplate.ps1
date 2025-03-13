function Get-EmbeddedHTMLTemplate {
    <#
    .SYNOPSIS
        Returns an embedded HTML template for migration reports.
    
    .DESCRIPTION
        Provides a fallback HTML template for generating migration reports when
        an external template file is not available. Ensures reports can be
        generated even without the external template file.
    
    .EXAMPLE
        $template = Get-EmbeddedHTMLTemplate
    
    .OUTPUTS
        [string] HTML template content as a string.
    #>
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
        body { font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: #fff; padding: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        h1, h2, h3 { color: #0078d4; }
        h1 { border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
        .summary { background-color: #f0f8ff; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 5px solid #0078d4; }
        .status-ready { background-color: #dff0d8; color: #3c763d; padding: 5px 10px; border-radius: 3px; font-weight: bold; }
        .status-warning { background-color: #fcf8e3; color: #8a6d3b; padding: 5px 10px; border-radius: 3px; font-weight: bold; }
        .status-failed { background-color: #f2dede; color: #a94442; padding: 5px 10px; border-radius: 3px; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th, td { text-align: left; padding: 12px 15px; border-bottom: 1px solid #ddd; }
        th { background-color: #0078d4; color: white; }
        tr:hover { background-color: #f5f5f5; }
        .mailbox-details { margin-top: 20px; margin-bottom: 30px; border: 1px solid #ddd; border-radius: 5px; overflow: hidden; }
        .mailbox-details h3 { margin: 0; padding: 15px; background-color: #e9ecef; border-bottom: 1px solid #ddd; }
        .mailbox-details-content { padding: 15px; }
        .detail-row { display: flex; margin-bottom: 8px; border-bottom: 1px solid #eee; padding-bottom: 8px; }
        .detail-label { width: 30%; font-weight: bold; }
        .detail-value { width: 70%; }
        .error-list, .warning-list { margin-top: 10px; padding-left: 20px; }
        .error-list li { color: #a94442; }
        .warning-list li { color: #8a6d3b; }
        .badge { display: inline-block; min-width: 10px; padding: 3px 7px; font-size: 12px; font-weight: 700; line-height: 1; color: #fff; text-align: center; white-space: nowrap; vertical-align: middle; background-color: #777; border-radius: 10px; }
        .badge-success { background-color: #5cb85c; }
        .badge-warning { background-color: #f0ad4e; }
        .badge-danger { background-color: #d9534f; }
        .timestamp { color: #777; font-style: italic; margin-top: 30px; text-align: center; }
        .error-code { display: inline-block; background-color: #f8d7da; color: #721c24; padding: 0px 5px; border-radius: 3px; font-family: monospace; margin-right: 5px; }
        .action-required { background-color: #f8d7da; color: #721c24; padding: 10px; border-radius: 5px; margin-top: 20px; border-left: 5px solid #721c24; }
        .tabs { display: flex; margin-top: 20px; border-bottom: 1px solid #ddd; }
        .tab { padding: 10px 15px; cursor: pointer; background-color: #f1f1f1; border: 1px solid #ddd; border-bottom: none; margin-right: 5px; border-top-left-radius: 5px; border-top-right-radius: 5px; position: relative; top: 1px; }
        .tab.active { background-color: white; border-bottom: 1px solid white; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        .category { margin-top: 15px; padding: 10px; background-color: #f8f9fa; border-radius: 3px; border-left: 3px solid #0078d4; }
        .category h4 { margin-top: 0; color: #0078d4; }
        .yes-value { color: #3c763d; font-weight: bold; }
        .no-value { color: #a94442; font-weight: bold; }
        .print-button { background-color: #0078d4; color: white; border: none; padding: 8px 15px; border-radius: 4px; cursor: pointer; margin-bottom: 15px; }
        .print-button:hover { background-color: #005a9e; }
        @media print { body { background-color: white; } .container { box-shadow: none; } .tabs, .tab, .print-button { display: none; } .tab-content { display: block; } }
    </style>
</head>
<body>
    <div class="container">
        <button class="print-button" onclick="window.print()">Print Report</button>
        <h1>Exchange Online Migration Report</h1>
        
        <div class="summary">
            <h2>Migration Batch: {{BatchName}}</h2>
            <p><strong>Report Generated:</strong> {{ReportDate}}</p>
            <p><strong>Total Mailboxes:</strong> {{TotalCount}}</p>
            <p><strong>Ready for Migration:</strong> <span class="badge badge-success">{{ReadyCount}}</span></p>
            <p><strong>Warnings:</strong> <span class="badge badge-warning">{{WarningCount}}</span></p>
            <p><strong>Failed:</strong> <span class="badge badge-danger">{{FailedCount}}</span></p>
            <p><strong>Script Version:</strong> {{ScriptVersion}}</p>
            {{BatchDetails}}
        </div>
        
        <div class="tabs">
            <div class="tab active" onclick="openTab(event, 'summary-tab')">Summary</div>
            <div class="tab" onclick="openTab(event, 'details-tab')">Detailed Results</div>
            <div class="tab" onclick="openTab(event, 'migration-tab')">Migration Guidance</div>
        </div>
        
        <div id="summary-tab" class="tab-content active">
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
        </div>
        
        <div id="details-tab" class="tab-content">
            <h2>Detailed Results</h2>
            {{DetailedResults}}
        </div>
        
        <div id="migration-tab" class="tab-content">
            <h2>Migration Guidance</h2>
            
            <div class="category">
                <h4>Critical Issues</h4>
                <p>The following issues must be resolved before migration:</p>
                <ul>
                    {{CriticalIssues}}
                </ul>
            </div>
            
            <div class="category">
                <h4>Performance Considerations</h4>
                <p>The following issues may impact migration performance:</p>
                <ul>
                    {{PerformanceIssues}}
                </ul>
            </div>
            
            <div class="category">
                <h4>Special Mailbox Types</h4>
                <p>The following special mailbox types require specific handling:</p>
                <ul>
                    {{SpecialMailboxes}}
                </ul>
            </div>
            
            <div class="category">
                <h4>Post-Migration Tasks</h4>
                <p>The following tasks should be performed after migration:</p>
                <ul>
                    {{PostMigrationTasks}}
                </ul>
            </div>
        </div>
        
        {{ActionRequired}}
        
        <p class="timestamp">Generated by Exchange Online Migration Script v{{ScriptVersion}}</p>
    </div>
    
    <script>
    function openTab(evt, tabName) {
        var i, tabContent, tabLinks;
        
        tabContent = document.getElementsByClassName("tab-content");
        for (i = 0; i < tabContent.length; i++) {
            tabContent[i].className = tabContent[i].className.replace(" active", "");
        }
        
        tabLinks = document.getElementsByClassName("tab");
        for (i = 0; i < tabLinks.length; i++) {
            tabLinks[i].className = tabLinks[i].className.replace(" active", "");
        }
        
        document.getElementById(tabName).className += " active";
        evt.currentTarget.className += " active";
    }
    </script>
</body>
</html>
'@
}
