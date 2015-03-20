##################################################################################################################
# FileSync.ps1
##################################################################################################################
# Script used to perform a file sync operation.
#
# !!!WARNING!!!
# Before deploying this script, you should sign it with an Authenticode signature.
# This will prevent tampering and allow you to use the more secure RemoteSigned execution policy on clients.
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

# XML result document that gets saved to a server
[xml]$jobResultsXml = New-Object System.Xml.XmlDocument
$jobResultsXml.LoadXml("<?xml version=`"1.0`" encoding=`"utf-8`"?><jobs></jobs>")

# Associative array of job results
$jobResultsArray = @{}

# Detect if this instance is interactive or not
$interactive = !((gwmi win32_process -Filter "ProcessID=$PID" | ? { $_.ProcessName -eq "powershell.exe" }).commandline -match "-NonInteractive")

# This function pauses script execution until a key is pressed
Function Pause
{
    Param (
        [Parameter(Mandatory=$true)] [string] $prompt
    )

    if ($interactive)
    {
        Write-Host -NoNewline $prompt
        [void][System.Console]::ReadKey($TRUE)
        Write-Host ""
    }
}

# Reads the robocopy log and parses the last relevant line for progress information
Function Parse-RoboCopyLog
{
    Param (
        [Parameter(Mandatory=$true)] [string] $logFile
    )
    
    try
    {
        $logLine = (Get-Content $logFile -ErrorAction SilentlyContinue | Select-Object -last 1)
    }
    catch
    {
        return
    }

    # Directory scan match
    if ($logLine -Match '[\s]*([\d]+)[\s]*\\\\([\w\W]+)')
    {
        $script:lastActivity = "Scanning directory..."
        $script:lastStatus = "\\" + $matches[2]
        Write-Progress -Id 2 -ParentId 1 -Activity $script:lastActivity -Status $script:lastStatus
        return
    }
    
    # New file match
    if (($logLine -Match 'New File[\s]+([\d.\w ]+)[\s]+([\w\W]+)') -or
        ($logLine -Match 'Newer[\s]+([\d.\w ]+)[\s]+([\w\W]+)'))
    {
        $script:lastActivity = "Copying file..."
        $script:lastStatus = $matches[2]
        Write-Progress -Id 2 -ParentId 1 -Activity $script:lastActivity -Status $script:lastStatus -PercentComplete 0
        return
    }
    
    # Percent completed match
    if ($logLine -Match '([\d]{1,3})%')
    {
        if ($matches.Count -gt 1)
        {
            [int] $progress = [int] $matches[1]
            Write-Progress -Id 2 -ParentId 1 -Activity $script:lastActivity -Status $script:lastStatus -PercentComplete $progress
        }
        return
    }
}

# This function executes a robocopy job and saves the job result as XML
Function Execute-Job
{
    Param (
        [Parameter(Mandatory=$true)] [string] $jobName,
        [Parameter(Mandatory=$true)] [string] $jobArguments
    )
    
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = "robocopy.exe"
    $processInfo.RedirectStandardError = $true
    $processInfo.RedirectStandardOutput = $true
    $processInfo.UseShellExecute = $false
    $processInfo.Arguments = "$jobArguments /log:`"$logFolder\$jobName.txt`""
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    $process.Start() | Out-Null

    do
    {
        if ($interactive)
        {
            Parse-RoboCopyLog "$logFolder\$jobName.txt"
        }
        Start-Sleep -m 1
    }
    while (Get-Process -Id $process.Id -ErrorAction SilentlyContinue | select -Property Responding)
    
    Write-Progress -Id 2 -ParentId 1 -Activity "Done" -Status "Done" -Completed
    
    $stopwatch.Stop()
    
    $elapsed = $stopwatch.Elapsed
    $return = Parse-RoboCopyReturnCode $process.ExitCode
    
    # Add the job result to the XML log file
    Add-JobResultXml $jobName $return $jobArguments $elapsed
    
    # Add the job result to the array of completed jobs
    $jobResultsArray[$jobName] = $return
}

# Adds a new row to the XML log file
Function Add-JobResultXml
{
    Param (
        [Parameter(Mandatory=$true)] [string] $jobName,
        [Parameter(Mandatory=$true)] [string] $jobReturn,
        [Parameter(Mandatory=$true)] [string] $jobArguments,
        [Parameter(Mandatory=$true)] [string] $jobTime
    )
    
    $resultXml = $jobResultsXml.CreateElement("result")
    
    $resultXmlName = $jobResultsXml.CreateAttribute("name")
    $resultXmlName.Value = $jobName
    
    $resultXmlCmd = $jobResultsXml.CreateAttribute("cmd")
    $resultXmlCmd.Value = $jobArguments
    
    $resultXmlTime = $jobResultsXml.CreateAttribute("time")
    $resultXmlTime.Value = $jobTime
    
    $resultXmlText = $jobResultsXml.CreateTextNode($jobReturn)
    
    $resultXml.Attributes.Append($resultXmlCmd) | Out-Null # pipe to Out-Null to suppress messages to the console
    $resultXml.Attributes.Append($resultXmlTime) | Out-Null
    $resultXml.Attributes.Append($resultXmlName) | Out-Null
    $resultXml.AppendChild($resultXmlText) | Out-Null
    
    $jobResultsXml.LastChild.AppendChild($resultXml) | Out-Null
}

# Saves the job output to an XML file on a remote server
Function Save-JobResultXml
{
    Param (
        [Parameter(Mandatory=$true)] [string] $remoteLogFolder,
        [Parameter(Mandatory=$true)] [string] $scriptRunTime,
        [Parameter(Mandatory=$true)] [string] $scriptStartTime,
        [Parameter(Mandatory=$true)] [string] $scriptFinishTime
    )

    $totalRunXml = $jobResultsXml.CreateAttribute("duration")
    $totalRunXml.Value = $scriptRunTime
    $jobResultsXml.SelectSingleNode("/jobs").Attributes.Append($totalRunXml) | Out-Null
    
    $startedXml = $jobResultsXml.CreateAttribute("started")
    $startedXml.Value = $scriptStartTime
    $jobResultsXml.SelectSingleNode("/jobs").Attributes.Append($startedXml) | Out-Null
    
    $finishedXml = $jobResultsXml.CreateAttribute("finished")
    $finishedXml.Value = $scriptFinishTime
    $jobResultsXml.SelectSingleNode("/jobs").Attributes.Append($finishedXml) | Out-Null
    
    # Write output to remote location
    $hostName = $env:COMPUTERNAME
    $jobResultsXml.Save("$remoteLogFolder\$hostName.xml")
}

# Parses the RoboCopy return code to determine status. Returns a string message.
# RoboCopy returns its status as a bit flag
Function Parse-RoboCopyReturnCode
{
    Param (
        [Parameter(Mandatory=$true)] [int] $returnCode
    )
    
    if ($returnCode -eq 0)
    {
        return "OK (NoChange)"
    }
    
    $retStr = ""
    
    if ($returnCode -band 1)
    {
        $retStr += "OK (CopyChanges) "
    }
    
    if ($returnCode -band 2)
    {
        $retStr += "Xtra "
    }
    
    if ($returnCode -band 4)
    {
        $retStr += "MISMATCHES "
    }
    
    if ($returnCode -band 8)
    {
        $retStr += "FAIL "
    }
    
    if ($returnCode -band 16)
    {
        $retStr += "FATALERROR "
    }
    
    return $retStr
}

# Starts running sync jobs specified in the $configFileLocation parameter
Function Start-SyncJobs
{
    Param (
        [Parameter(Mandatory=$true)] [string] $configFileLocation
    )

    $config = [xml](Get-Content $configFileLocation)

    $rootNode = $config.SelectSingleNode("/syncConfig")
    $logFolder = $config.SelectSingleNode("/syncConfig").getAttribute("localLogFolder")
    $retries = "/r:" + $rootNode.getAttribute("retries")
    $retryWaitTime = "/w:" + $rootNode.getAttribute("retryWaitTime")
    $global:syncBatchName = $rootNode.getAttribute("syncBatchName")

    # Create the path to the sync log folder and hide the log folder
    $logFolderObj = New-Item -ItemType Directory -Force -Path $logFolder
    Set-ItemProperty $logFolderObj -Name Attributes -Value "Hidden"

    # Get all the fileSync jobs
    $syncJobs = $config.SelectNodes("/syncConfig/fileSync")

    $scriptStarted = (Get-Date -Format "G")    
    $scriptTimer = [Diagnostics.Stopwatch]::StartNew()
    
    if ($interactive)
    {
        Clear-Host
        Write-Host "Beginning file sync jobs. This might take a while, so please be patient."
    }
    
    $currentJobIndex = 0
    
    # Run each job!
    foreach ($job in $syncJobs)
    {
        # Job parameters
        $jobName = $job.getAttribute("jobName")
        $remoteRoot = $job.getAttribute("remoteRoot")
        $localRoot = $job.getAttribute("localRoot")
        $folderMode = ""
        $exclusions = ""
    
        # Detect if we want to include empty folders (/e) or exclude them (/s)
        if ($job.getAttribute("includeEmptyFolders") -eq "true")
        {
            $folderMode = "/e"
        }
        else
        {
            $folderMode = "/s"
        }
    
        # Check for exclusion nodes
        if ($job.HasChildNodes)
        {
            $exclusions = "/xd "
            
            foreach ($exclusion in $job.ChildNodes)
            {
                $exclusions += $exclusion.InnerXML + " "
            }
        }
        
        # Build the argument list and kick it off
        $arguments = "`"$remoteRoot`" `"$localRoot`" /purge $folderMode $exclusions $retries $retryWaitTime"
        
        if ($interactive)
        {
            Write-Progress -Id 1 -Activity ("Synchronizing " + $global:syncBatchName + "...") -Status "Executing job `"$jobName`"" -PercentComplete ((++$currentJobIndex / $syncJobs.Count) * 100)
            $script:lastActivity = "Initializing..."
            $script:lastStatus = "Starting up"
        }
        
        Execute-Job $jobName $arguments
    }
    
    Write-Progress -Id 1 -Activity "Done" -Status "Done" -Complete
    
    # Performance timing
    $scriptTimer.Stop()
    $scriptElapsed = $scriptTimer.Elapsed
    
    if ($interactive)
    {
        Write-Host ""
        Write-Host "Finished file sync jobs ($scriptElapsed)"
        Write-Host ""
        Write-Host "Job Summary:"
        $jobResultsArray.GetEnumerator() | Sort-Object Name | Format-Table -AutoSize # Print job results as a formatted table
        Pause "Press any key to exit."
    }

    # Save the XML report to a remote server location
    Save-JobResultXml $config.SelectSingleNode("/syncConfig").getAttribute("remoteLogFolder") $scriptElapsed $scriptStarted (Get-Date -Format "G")
}

# Begin!

Start-SyncJobs $configFilePath

