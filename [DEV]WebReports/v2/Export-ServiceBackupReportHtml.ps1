param (
	[Parameter(Mandatory=$false)]
	[ValidateSet("AzureAD","Exchange","Intune","Teams","SharePoint","OneDrive","SecurityCompliance")]
	[string]$Service,
	[Parameter(Mandatory=$false)]
	[ValidateSet("Full backup","Services backup","Components backup")]
	[string]$BackupType,
	[Parameter(Mandatory=$false)]
	[switch]$DebugEnabled
)

if($DebugEnabled) {
	$Debug = $true
} else {
	$Debug = $false
}

#Initialize variables
$globalSuccess = $null
$successfulComponents = @()
$failedComponents = @()
$interruptedComponents = @()
$successPercentage = $null
$failurePercentage = $null
$interruptedPercentage = $null
$newModuleUpdateAvailable = $null
$logFilePath = $null
$logContent = $null
$isSuccessList = $false
$isFailList = $false
$isInterruptedList = $false
$Title = $null

#Script location retrieved automatically
$ScriptLocation = $PSScriptRoot
#If no script location can be retrieved automatically, script location will be current location
if(!($PSScriptRoot -match "\\")) { $ScriptLocation = Get-Location }

#Script name retrieved automatically
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0]
#If no script name can be retrieved automatically, script name will be defined as below
if(!($MyInvocation.MyCommand.Name -match "\w")) { $ScriptName = "Export-ServiceBackupReportHtml" } 

#Configuration files directory
$DSCExport = $ScriptLocation

Set-Location $DSCExport

$FormattedDateForDirectoryFormat = Get-Date -Format "yyyyMMdd"
$FormattedDate = $(Get-Date).ToString("dd_MM_yyyy_HH-mm-ss")
$dateTime = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"

#Transcript file
$TranscriptFile = "$DSCExport\Transcript_$ScriptName`_$FormattedDate.log"

Start-Transcript -Path $TranscriptFile -Force

if($Service -and $BackupType -ne "Services backup") {
	"[ERROR] : -Service parameter can only be specified using -BackupType 'Services backup' parameter"
	Stop-Transcript
	exit
	
}

if($BackupType -eq "Services backup" -and ([string]::IsNullOrEmpty($Service))) {
	"[ERROR] : -BackupType 'Services backup' parameter can only be used with -Service parameter"
	Stop-Transcript
	exit
}

if($Debug) { "BackupType: {0}" -f $BackupType }

#Define the path to the log file
switch($BackupType) {
	"Full backup" { 
	if($Debug) { "Retrieve Backup_FullBackup last txt file based on last write time" }	
		try { 
			$logFilePath = Get-ChildItem $DSCExport -ErrorAction Stop | ?{$_.name.StartsWith("Backup_FullBackup") -and $_.name.EndsWith(".txt")} | sort -Descending LastWriteTime | select -First 1 
			"[OK] : Last full backup txt file retrieved : {0}" -f $logFilePath
		} catch {
			"[ERROR] : Fail to retrieve last full backup txt file : {0}" -f $_.Exception[0].Message
			Stop-Transcript
			exit
		}
	}
	"Services backup" { 
	if($Debug) { "Retrieve Backup_$Service last txt file based on last write time" }	
		try {
			$logFilePath = Get-ChildItem $DSCExport -ErrorAction Stop | ?{$_.name.StartsWith("Backup_$Service") -and $_.name.EndsWith(".txt")} | sort -Descending LastWriteTime | select -First 1 
			"[OK] : Last service {0} backup txt file retrieved : {1}" -f $Service,$logFilePath
		} catch {
			"[ERROR] : Fail to retrieve last service {0} backup txt file : {1}" -f $Service,$_.Exception[0].Message
			Stop-Transcript
			exit			
		}
	}
	"Components backup" {
	if($Debug) { "Retrieve Backup_customComponents last txt file based on last write time" }
		try {
			$logFilePath = Get-ChildItem $DSCExport -ErrorAction Stop | ?{$_.name.StartsWith("Backup_customComponents") -and $_.name.EndsWith(".txt")} | sort -Descending LastWriteTime | select -First 1 
			"[OK] : Last components backup txt file retrieved : {0}" -f $logFilePath
		} catch {
			"[ERROR] : Fail to retrieve last components backup txt file : {0}" -f $_.Exception[0].Message
			Stop-Transcript
			exit				
		}
	}
}
if($Debug) { "Target backup txt file : {0}\{1}" -f $DSCExport,$logFilePath }

#Read the log file content
if(!([string]::IsNullOrEmpty($logFilePath))) {
	try {
		$logContent = Get-Content -Path "$DSCExport\$logFilePath" -Encoding Default	
		"[OK] : Content retrieved from file {0}\{1}" -f $DSCExport,$logFilePath
	} catch {
		"[ERROR] : Fail to retrieve content from file {0}\{1} : {2}" -f $DSCExport,$logFilePath,$_.Exception[0].Message
		Stop-Transcript
		exit		
	}
} else {
	"[ERROR] : Fail to retrieve data from backup txt file : {0}" -f $logFilePath
	Stop-Transcript
	exit		
}

#Parse the file content
$backupStart = (($logContent | Select-String "backup start") -split(": "))[1]
if($Debug) { "backupStart: {0}" -f $backupStart }
$backupEnd = (($logContent | Select-String "backup end") -split(": "))[1]
if($Debug) { "backupEnd: {0}" -f $backupEnd }
$backupDuration = (($logContent | Select-String "backup duration") -split(": "))[1]
if($Debug) { "backupDuration: {0}" -f $backupDuration }

foreach ($line in $logContent) {	
	if($Debug) { "Processing line: $line" }

    if ($line -match "Global Backup Job Status: (\w+)") {
		if($Debug) { '$line -match "Global Backup Job Status: (\w+)"' }
        $globalSuccess = $matches[1]
    }
    elseif ($line -match "Successfull backup") {
		if($Debug) { '$line -match "Successfull backup"' }
        $isSuccessList = $true
    }
    elseif ($line -match "Failed backup") {
		if($Debug) { '$line -match "Failed backup"' }
        $isFailList = $true
    }
	elseif ($line -match "Interrupted backup") {
		if($Debug) { '$line -match "Interrupted backup"' }
        $isInterruptedList = $true
    }
	
	elseif ($line -match "Global Success Percentage: \d{1,3}%") {
		if($Debug) { '$line -match "Global Success Percentage: \d{1,3}%"' }
        $successPercentage = $matches[0]
		$successPercentage = ($successPercentage -split(":"))[1]
    }
    elseif ($line -match "Global Failure Percentage: \d{1,3}%") {
		if($Debug) { '$line -match "Global Failure Percentage: \d{1,3}%"' }
        $failurePercentage = $matches[0]
		$failurePercentage = ($failurePercentage -split(":"))[1]
    }
	elseif ($line -match "Global Interrupted Percentage: \d{1,3}%") {
		if($Debug) { '$line -match "Global Interrupted Percentage: \d{1,3}%"' }
        $interruptedPercentage = $matches[0]
		$interruptedPercentage = ($interruptedPercentage -split(":"))[1]
    }
	
	#Check for new module update availability
    elseif ($line -match "New PowerShell Module Update Available: (\w+)") {
		if($Debug) { '$line -match "New PowerShell Module Update Available: (\w+)"' }
        $newModuleUpdateAvailable = $matches[1]
    }
    elseif ($line -match "^\s*(\w+)\s*$") {
		if($Debug) { '$line -match "^\s*(\w+)\s*$"' }
        if ($isSuccessList) {
            $successfulComponents += $matches[1]
        } if ($isFailList) {
            $failedComponents += $matches[1]
        } if ($isInterruptedList) {
			$interruptedComponents += $matches[1]
		}
    }
}

if($BackupType -eq "Full backup" -or $BackupType -eq "Components backup") { 
	$Title = "$BackupType report" 
} elseif($BackupType -eq "Services backup") {
	$Title = "$Service backup report"
}
if($Debug) { "BackupType : {0} - Title : {1}" -f $BackupType,$Title }

if($globalSuccess -eq "SUCCESS") {
	$Class = '<span class="success">'
} elseif($globalSuccess -eq "WARNING" -or $globalSuccess -eq "INTERRUPTED") {
	$Class = '<span class="warning">'
} elseif($globalSuccess -eq "FAILURE") {
	$Class = '<span class="failed">'
}

if($Debug) { "Generate HTML report" }
$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>$Title</title>
    <style>
        body { font-family: ui-sans-serif, -apple-system, system-ui, Segoe UI, Helvetica, Apple Color Emoji, Arial, sans-serif, Segoe UI Emoji, Segoe UI Symbol; font-size: 14px; margin: 20px; padding: 0px; color: #333; }
        h2 { color: #333333; margin-top: 20px; }
		h3 { color: #333333; margin-top: 20px; }
        table { align-items: center; width: 20%; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.2);}
        td { padding: 12px; border-bottom: 1px solid #ddd; }
        th { background-color: #f0f0f0; padding: 12px; text-align: left; }
		.status { display: flex; align-items: left;gap: 8px; }
		.success { color: green; font-weight: bold; }
		.warning { color: orange; font-weight: bold; }
		.failed { color: red; font-weight: bold; }
    </style>
</head>
<body>
	
    <h2>$Title</h2>
    <p><strong>📅 Backup Date:</strong> $dateTime</p>
    <p><strong>⏱ Start Time:</strong> $backupStart</p>
    <p><strong>⏱ End Time:</strong> $backupEnd</p>
    <p><strong>⏳ Duration:</strong> $backupDuration</p>
    <p><strong>💾 Global Backup Status:</strong>$Class $globalSuccess</span></p>
    <p><strong>🔵 Success Rate:</strong> $successPercentage 📈</p>
    <p><strong>🔴 Failure Rate:</strong> $failurePercentage 📉</p>
	<p><strong>⚫ Interrupted Rate:</strong> $interruptedPercentage 🚧</p>
	
    <h3>Backup Summary</h3>
    <table>
        <tr>
            <th>Status</th>
            <th>Items</th>
        </tr>
        <tr>
            <td class="status"><span class="success-icon">✅</span> Successful</td>
            <td>$(($successfulComponents.Count) -join ", ")</td>
        </tr>
        <tr>
            <td class="status"><span class="failed-icon">❌</span> Failed</td>
            <td>$(($failedComponents.Count) -join ", ")</td>
        </tr>
		<tr>
            <td class="status"><span class="interrupted-icon">🔌</span> Interrupted</td>
            <td>$(($interruptedComponents.Count) -join ", ")</td>
        </tr>
    </table><br />
	
	<h3>🔵</span> Successful Backups ($($successfulComponents.Count)) </h3>
    <table style="width: 50% ;"><tr>$($nbComp = 0;foreach ($successfulComponent in $successfulComponents) { $nbComp++;if($nbComp -ge 8) { "</tr><tr>";$nbComp = 0};"<td style='width: 40% ;text-align:left ;'>$successfulComponent</td>" }) </tr></table>
	
    <h3>🔴</span> Failed Backups ($($failedComponents.Count)) </h3>
	<table style="width: 50% ;"><tr>$($nbComp = 0;foreach ($failedComponent in $failedComponents) { $nbComp++;if($nbComp -ge 8) { "</tr><tr>";$nbComp = 0};"<td style='width: 40% ;text-align:left ;'>$failedComponent</td>" }) </tr></table>

	<h3>⚫</span> Interrupted Backups ($($interruptedComponents.Count)) </h3>
	<table style="width: 50% ;"><tr>$($nbComp = 0;foreach ($interruptedComponent in $interruptedComponents) { $nbComp++;if($nbComp -ge 8) { "</tr><tr>";$nbComp = 0};"<td style='width: 40% ;text-align:left ;'>$interruptedComponent</td>" }) </tr></table>
	
	<br /><p><strong>🚀 Action Required:</strong> Consider investigating the failed backup items.</p>
    $(if($newModuleUpdateAvailable -eq "Yes") { '<p><strong>🔥 New PowerShell Module Update Available !</strong></p>' })
	
</body>
</html>
"@	

#Set HTML name and destination
switch($BackupType) {
	"Full backup" { $htmlReportFilePath = "$DSCExport\Backup_FullBackup_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html" }
	"Services backup" { $htmlReportFilePath = "$DSCExport\Backup_$Service`_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html" }
	"Components backup" { $htmlReportFilePath = "$DSCExport\Backup_customComponents_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html" }	
}

#Export HTML file
try {
	$htmlReport | Out-File -FilePath $htmlReportFilePath -Encoding Default -Force -ErrorAction Stop
	"[OK] : Html report generated to destination {0}" -f $htmlReportFilePath
} catch {
	"[ERROR] : Fail to generate html report {0} : {1}" -f $htmlReportFilePath,$_.Exception[0].Message
}

#Open the HTML report in the default browser
#Start-Process $htmlReportFilePath

Stop-Transcript