##################################################################################################################
# SyncLogReporting.ps1
##################################################################################################################
# Script used to gather logs generated by FileSync.ps1 and send e-mail reports.
#
# !!!WARNING!!!
# Before deploying this script, you should sign it with an Authenticode signature.
# This will prevent tampering and allow you to use the more secure RemoteSigned execution policy.
#
# You can sign this script with the following PowerShell commands. Please note that you need to have a code
# signing cert from your enterprise CA installed in your Personal cert store.
#
# $cert = (dir cert:currentuser\my\ -CodeSigningCert)
# Set-AuthenticodeSignature $script $cert -TimeStampServer "http://timestamp.verisign.com/scripts/timstamp.dll"
##################################################################################################################
# For more information on this script, refer to the GitHub repository at 
#    https://github.com/JohnCardenas/FileSyncScript
##################################################################################################################

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true,Position=1)] [string] $configFilePath
)

Import-Module ActiveDirectory

# Array of computers that did not synchronize
$global:noSyncLog = New-Object System.Collections.ArrayList

# Hash table of sync logs from member computers
$global:syncLogs = @{}

# Hash table of job names
$global:jobNames = @{}

# Reads the sync config file for paths and job names
Function Read-ConfigFile
{
    Param(
        [Parameter(Mandatory=$true)] [string] $configFilePath
    )

    [xml]$config = Get-Content $configFilePath
    $syncJobs = $config.SelectNodes("/syncConfig/fileSync")
    
    foreach ($job in $syncJobs)
    {
        $global:jobNames[$job.getAttribute("jobName")] = $job.getAttribute("displayName")
    }
    
    $syncConfig = $config.SelectSingleNode("/syncConfig")
    
    $global:remoteLogFolder    = $syncConfig.getAttribute("remoteLogFolder")
    $global:computerOU         = $syncConfig.getAttribute("computerOU")
    $global:reportEmailServer  = $syncConfig.getAttribute("reportEmailServer")
    $global:reportEmailFrom    = $syncConfig.getAttribute("reportEmailFrom")
    $global:reportEmailTo      = $syncConfig.getAttribute("reportEmailTo")
    $global:reportEmailSubject = $syncConfig.getAttribute("reportEmailSubject")
}

# Reads the specified sync log into variables
Function Read-SyncLog
{
    Param(
        [Parameter(Mandatory=$true)] [System.IO.FileInfo] $logFile
    )
    
    # Remove the computer from the noSyncLog list since we found one for this computer
    if ($global:noSyncLog -Contains $logFile.BaseName)
    {
        $global:noSyncLog.Remove($logFile.BaseName)
    }
    
    # Read the log in
    [xml]$syncLog = Get-Content $logFile.FullName
    $syncJobs = $syncLog.SelectNodes("/jobs/result")
    
    # Job results
    $jobResults = New-Object System.Collections.HashTable
    
    foreach ($job in $syncJobs)
    {
        $jobResults[$job.Name] = $job."#text"
    }
    
    # Push the results
    $global:syncLogs[$logFile.BaseName] = $jobResults
}

# Gets a list of all sync logs in the log directory
Function Get-SyncLogs
{
    $logList = Get-ChildItem $global:remoteLogFolder -Filter "*.xml"
    
    foreach ($log in $logList)
    {
        Read-SyncLog $log
    }
}

# Gets a list of all computers in AD
Function Get-Computers
{
    # Get the list of computers in AD
    $computerList = Get-ADComputer -Filter * -SearchBase $global:computerOU
    
    foreach ($computer in $computerList)
    {
        $global:noSyncLog.Add($computer.Name) | Out-Null
    }
    
    $global:noSyncLog.Sort()
}

# Formats the sync log results into an HTML table
Function Format-SyncLogs
{
    $table = "<table border='1'><thead><tr>"
    $table += "<td align='center' valign='middle'><b>Job Name</b></td>"
    
    # Sort computer names alphabetically
    $sortedComputerList = New-Object System.Collections.ArrayList
    foreach ($computer in $global:syncLogs.Keys)
    {
        $sortedComputerList.Add($computer) | Out-Null
    }
    $sortedComputerList.Sort()
    
    # Add computer names to header
    foreach ($computer in $sortedComputerList)
    {
        $table += "<td align='center' valign='middle'><b>" + $computer + "</b></td>"
    }
    
    # Close headers
    $table += "</tr></thead>"
    $table += "<tbody>"
    
    # Sort the job list alphabetically
    $sortedJobList = $global:jobNames.GetEnumerator() | Sort-Object Value

    foreach ($job in $sortedJobList)
    {
        $table += "<tr>"
        $table += "<td>" + $job.Value + "</td>"
        
        foreach ($computer in $sortedComputerList)
        {
            $retVal = $global:syncLogs[$computer][$job.Name]
            
            if ($retVal.Contains("OK"))
            {
                $table += "<td valign='middle' align='center' bgcolor='#00FF00'>OK</td>"
            }
            elseif ($retVal.Contains("FAIL") -or $retVal.Contains("FATALERROR"))
            {
                $table += "<td valign='middle' align='center' bgcolor='#FF0000'>FAIL</td>"
            }
            else
            {
                $table += "<td valign='middle' align='center' bgcolor='#FFFF00'>$retVal</td>"
            }
        }
        
        $table += "</tr>"
    }

    $table += "</tbody></table>"
    
    return $table
}

# Sends an HTML report e-mail
Function Send-ReportEmail
{
    $reportBody = "<h1>Field Unit Engineering Images Sync Report</h1>"
    $reportBody += Format-SyncLogs
    
    if ($global:noSyncLog.Count -gt 0)
    {
        $reportBody += "<p>The computers below did not report their status. Please confirm that they have been connected to the network and run the Update Images script if needed.</p>"
        $reportBody += "<ul>"
        
        foreach ($computer in $global:noSyncLog)
        {
            $reportBody += "<li>$computer</li>"
        }
        
        $reportBody += "</ul>"
    }
    
    # Build and send the report e-mail
    $date = Get-Date -Format "yyyy-MM-dd"
    $report = New-Object Net.Mail.MailMessage
    $report.From = $global:reportEmailFrom
    $report.Subject = $global:reportEmailSubject + " ($date)"
    $report.To.Add($global:reportEmailTo)
    $report.IsBodyHtml = $TRUE
    $report.Body = $reportBody
    
    $smtp = New-Object Net.Mail.SmtpClient($global:reportEmailServer)
    $smtp.Send($report)
}

# Deletes logs from the log folder
Function Cleanup-LogFolder
{
    Remove-Item $global:remoteLogFolder\* -Recurse
}

# Begin!
Read-ConfigFile $configFilePath
Get-Computers
Get-SyncLogs
Send-ReportEmail
Cleanup-LogFolder

