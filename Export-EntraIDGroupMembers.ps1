<#
.SYNOPSIS
    Exports members of Entra ID security groups to CSV and generates an HTML report.

.DESCRIPTION
    This script connects to Microsoft Graph, reads a CSV file containing Entra ID security group display names,
    and exports all members of those groups to a new CSV file. It also generates a comprehensive HTML report
    documenting the export process.

.PARAMETER InputCsvPath
    Path to the CSV file containing group display names (required column: DisplayName)

.PARAMETER OutputCsvPath
    Path for the output CSV file (default: ".\EntraID_GroupMembers_Export.csv")

.PARAMETER ReportPath
    Path for the HTML report file (default: ".\EntraID_Export_Report.html")

.EXAMPLE
    .\Export-EntraIDGroupMembers.ps1 -InputCsvPath ".\groups.csv"
    
.EXAMPLE
    .\Export-EntraIDGroupMembers.ps1 -InputCsvPath ".\groups.csv" -OutputCsvPath ".\members.csv" -ReportPath ".\report.html"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$InputCsvPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputCsvPath = ".\EntraID_GroupMembers_Export.csv",
    
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = ".\EntraID_Export_Report.html"
)

# Initialize variables
$StartTime = Get-Date
$ExportedMembers = @()
$ProcessingResults = @()
$TotalGroups = 0
$SuccessfulGroups = 0
$FailedGroups = 0
$TotalMembers = 0

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to generate HTML report
function Generate-HTMLReport {
    param(
        [array]$Results,
        [string]$OutputPath,
        [datetime]$StartTime,
        [datetime]$EndTime,
        [int]$TotalGroups,
        [int]$SuccessfulGroups,
        [int]$FailedGroups,
        [int]$TotalMembers
    )
    
    $Duration = $EndTime - $StartTime
    $ResultsHtml = ""
    
    foreach ($result in $Results) {
        $statusColor = if ($result.Status -eq "Success") { "#28a745" } else { "#dc3545" }
        $ResultsHtml += @"
        <tr>
            <td>$($result.GroupName)</td>
            <td>$($result.GroupId)</td>
            <td>$($result.MemberCount)</td>
            <td><span style="color: $statusColor; font-weight: bold;">$($result.Status)</span></td>
            <td>$($result.Message)</td>
        </tr>
"@
    }
    
    $HtmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Entra ID Group Members Export Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 3px solid #0078d4; }
        .header h1 { color: #0078d4; margin: 0; font-size: 2.5em; }
        .header p { color: #666; margin: 10px 0 0 0; font-size: 1.1em; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .summary-card { background: linear-gradient(135deg, #0078d4, #106ebe); color: white; padding: 20px; border-radius: 8px; text-align: center; }
        .summary-card h3 { margin: 0 0 10px 0; font-size: 2em; }
        .summary-card p { margin: 0; opacity: 0.9; }
        .section { margin-bottom: 30px; }
        .section h2 { color: #0078d4; border-bottom: 2px solid #e1e1e1; padding-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #f8f9fa; font-weight: 600; color: #495057; }
        tr:hover { background-color: #f8f9fa; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
        .footer { text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #e1e1e1; color: #666; }
        .info-box { background-color: #e7f3ff; border-left: 4px solid #0078d4; padding: 15px; margin: 15px 0; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Entra ID Group Members Export Report</h1>
            <p>Generated on $($EndTime.ToString("MMMM dd, yyyy 'at' HH:mm:ss"))</p>
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>$TotalGroups</h3>
                <p>Total Groups Processed</p>
            </div>
            <div class="summary-card">
                <h3>$SuccessfulGroups</h3>
                <p>Successful Exports</p>
            </div>
            <div class="summary-card">
                <h3>$FailedGroups</h3>
                <p>Failed Exports</p>
            </div>
            <div class="summary-card">
                <h3>$TotalMembers</h3>
                <p>Total Members Exported</p>
            </div>
        </div>
        
        <div class="section">
            <h2>Execution Details</h2>
            <div class="info-box">
                <strong>Start Time:</strong> $($StartTime.ToString("yyyy-MM-dd HH:mm:ss"))<br>
                <strong>End Time:</strong> $($EndTime.ToString("yyyy-MM-dd HH:mm:ss"))<br>
                <strong>Duration:</strong> $($Duration.ToString("hh\:mm\:ss"))<br>
                <strong>Input File:</strong> $InputCsvPath<br>
                <strong>Output File:</strong> $OutputCsvPath
            </div>
        </div>
        
        <div class="section">
            <h2>Group Processing Results</h2>
            <table>
                <thead>
                    <tr>
                        <th>Group Name</th>
                        <th>Group ID</th>
                        <th>Members Found</th>
                        <th>Status</th>
                        <th>Message</th>
                    </tr>
                </thead>
                <tbody>
                    $ResultsHtml
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p>Report generated by Export-EntraIDGroupMembers.ps1</p>
            <p>Microsoft Graph PowerShell SDK | Entra ID Group Export Tool</p>
        </div>
    </div>
</body>
</html>
"@
    
    $HtmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
}

# Main script execution
try {
    Write-ColorOutput "=== Entra ID Group Members Export Tool ===" "Cyan"
    Write-ColorOutput "Start Time: $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" "Gray"
    Write-ColorOutput ""
    
    # Check if required modules are installed
    Write-ColorOutput "Checking required modules..." "Yellow"
    
    $RequiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Groups", "Microsoft.Graph.Users")
    foreach ($Module in $RequiredModules) {
        if (!(Get-Module -ListAvailable -Name $Module)) {
            Write-ColorOutput "Installing module: $Module" "Yellow"
            Install-Module -Name $Module -Force -AllowClobber -Scope CurrentUser
        }
    }
    
    # Import modules
    Write-ColorOutput "Importing Microsoft Graph modules..." "Yellow"
    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Groups
    Import-Module Microsoft.Graph.Users
    
    # Connect to Microsoft Graph
    Write-ColorOutput "Connecting to Microsoft Graph..." "Yellow"
    Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All" -NoWelcome
    
    Write-ColorOutput "Successfully connected to Microsoft Graph!" "Green"
    Write-ColorOutput ""
    
    # Validate input file
    if (!(Test-Path $InputCsvPath)) {
        throw "Input CSV file not found: $InputCsvPath"
    }
    
    # Read input CSV
    Write-ColorOutput "Reading input CSV file: $InputCsvPath" "Yellow"
    $InputGroups = Import-Csv $InputCsvPath
    
    if (!$InputGroups -or $InputGroups.Count -eq 0) {
        throw "No groups found in input CSV file"
    }
    
    if (!($InputGroups | Get-Member -Name "DisplayName")) {
        throw "Input CSV must contain a 'DisplayName' column"
    }
    
    $TotalGroups = $InputGroups.Count
    Write-ColorOutput "Found $TotalGroups groups to process" "Green"
    Write-ColorOutput ""
    
    # Process each group
    $Counter = 1
    foreach ($InputGroup in $InputGroups) {
        $GroupName = $InputGroup.DisplayName.Trim()
        Write-ColorOutput "[$Counter/$TotalGroups] Processing group: $GroupName" "Cyan"
        
        try {
            # Find the group
            $Group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction Stop
            
            if (!$Group) {
                $ProcessingResults += [PSCustomObject]@{
                    GroupName = $GroupName
                    GroupId = "N/A"
                    MemberCount = 0
                    Status = "Failed"
                    Message = "Group not found"
                }
                $FailedGroups++
                Write-ColorOutput "  ❌ Group not found" "Red"
            }
            elseif ($Group.Count -gt 1) {
                $ProcessingResults += [PSCustomObject]@{
                    GroupName = $GroupName
                    GroupId = "Multiple"
                    MemberCount = 0
                    Status = "Failed"
                    Message = "Multiple groups found with same name"
                }
                $FailedGroups++
                Write-ColorOutput "  ❌ Multiple groups found with same name" "Red"
            }
            else {
                # Get group members
                $Members = Get-MgGroupMember -GroupId $Group.Id -All -ErrorAction Stop
                
                $MemberCount = 0
                foreach ($Member in $Members) {
                    try {
                        # Get detailed user information
                        $User = Get-MgUser -UserId $Member.Id -Property "DisplayName,UserPrincipalName,Mail,JobTitle,Department,CompanyName" -ErrorAction Stop
                        
                        $ExportedMembers += [PSCustomObject]@{
                            GroupName = $GroupName
                            GroupId = $Group.Id
                            MemberDisplayName = $User.DisplayName
                            MemberUPN = $User.UserPrincipalName
                            MemberEmail = $User.Mail
                            JobTitle = $User.JobTitle
                            Department = $User.Department
                            Company = $User.CompanyName
                            MemberId = $User.Id
                            ExportDate = $StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                        }
                        $MemberCount++
                        $TotalMembers++
                    }
                    catch {
                        Write-ColorOutput "    ⚠️ Could not retrieve details for member: $($Member.Id)" "Yellow"
                    }
                }
                
                $ProcessingResults += [PSCustomObject]@{
                    GroupName = $GroupName
                    GroupId = $Group.Id
                    MemberCount = $MemberCount
                    Status = "Success"
                    Message = "Successfully exported $MemberCount members"
                }
                $SuccessfulGroups++
                Write-ColorOutput "  ✅ Successfully exported $MemberCount members" "Green"
            }
        }
        catch {
            $ProcessingResults += [PSCustomObject]@{
                GroupName = $GroupName
                GroupId = "N/A"
                MemberCount = 0
                Status = "Failed"
                Message = $_.Exception.Message
            }
            $FailedGroups++
            Write-ColorOutput "  ❌ Error: $($_.Exception.Message)" "Red"
        }
        
        $Counter++
        Write-ColorOutput ""
    }
    
    # Export results to CSV
    if ($ExportedMembers.Count -gt 0) {
        Write-ColorOutput "Exporting $($ExportedMembers.Count) members to CSV: $OutputCsvPath" "Yellow"
        $ExportedMembers | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
        Write-ColorOutput "✅ CSV export completed successfully!" "Green"
    }
    else {
        Write-ColorOutput "⚠️ No members found to export" "Yellow"
    }
    
    # Generate HTML report
    $EndTime = Get-Date
    Write-ColorOutput "Generating HTML report: $ReportPath" "Yellow"
    Generate-HTMLReport -Results $ProcessingResults -OutputPath $ReportPath -StartTime $StartTime -EndTime $EndTime -TotalGroups $TotalGroups -SuccessfulGroups $SuccessfulGroups -FailedGroups $FailedGroups -TotalMembers $TotalMembers
    Write-ColorOutput "✅ HTML report generated successfully!" "Green"
    
    # Summary
    Write-ColorOutput ""
    Write-ColorOutput "=== EXPORT SUMMARY ===" "Cyan"
    Write-ColorOutput "Total Groups Processed: $TotalGroups" "White"
    Write-ColorOutput "Successful Exports: $SuccessfulGroups" "Green"
    Write-ColorOutput "Failed Exports: $FailedGroups" "Red"
    Write-ColorOutput "Total Members Exported: $TotalMembers" "White"
    Write-ColorOutput "Duration: $((Get-Date) - $StartTime)" "Gray"
    Write-ColorOutput ""
    Write-ColorOutput "Files Generated:" "Cyan"
    if ($ExportedMembers.Count -gt 0) {
        Write-ColorOutput "  📄 CSV Export: $OutputCsvPath" "White"
    }
    Write-ColorOutput "  📊 HTML Report: $ReportPath" "White"
    
}
catch {
    Write-ColorOutput "❌ Script execution failed: $($_.Exception.Message)" "Red"
    Write-ColorOutput "Stack Trace: $($_.ScriptStackTrace)" "Red"
    exit 1
}
finally {
    # Disconnect from Microsoft Graph
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-ColorOutput ""
        Write-ColorOutput "Disconnected from Microsoft Graph" "Gray"
    }
    catch {
        # Ignore disconnection errors
    }
}
