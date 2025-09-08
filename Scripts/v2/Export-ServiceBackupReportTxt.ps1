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
$componentsPlannedToBackup = @()
$totalComponents = $null
$newModuleUpdateAvailable = $false
$currentComponent = $null
$scriptException = $false
$Report = $null
$logFilePath = $null
$logContent = $null
$GlobalSuccessPercentage = $null
$GlobalFailurePercentage = $null
$GlobalInterruptedPercentage = $null
$successPercentage = $null
$failurePercentage = $null
$interruptedPercentage = $null

#Script location retrieved automatically
$ScriptLocation = $PSScriptRoot
#If no script location can be retrieved automatically, script location will be current location
if(!($PSScriptRoot -match "\\")) { $ScriptLocation = Get-Location }

#Script name retrieved automatically
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0]
#If no script name can be retrieved automatically, script name will be defined as below
if(!($MyInvocation.MyCommand.Name -match "\w")) { $ScriptName = "Export-ServiceBackupReportTxt" } 

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

#Define the path to the log file
switch($BackupType) {
	"Full backup" {
	if($Debug) { "Retrieve Transcript_FullBackup last file based on last write time" }		
		try {
			$logFilePath = Get-ChildItem $DSCExport -ErrorAction Stop | ?{$_.name.StartsWith("Transcript_FullBackup","CurrentCultureIgnoreCase")} | sort -Descending LastWriteTime | select -First 1
			"[OK] : Transcript full backup file retrieved : {0}" -f $logFilePath
		} catch {
			"[ERROR] : Fail to retrieve transcript full backup file : {0}" -f $_.Exception[0].Message
			Stop-Transcript
			exit
		}
	}
	"Services backup" { 
		if($Debug) { "Retrieve Transcript_$Service last file based on last write time" }
		try {
			$logFilePath = Get-ChildItem $DSCExport -ErrorAction Stop | ?{$_.name.StartsWith("Transcript_$Service","CurrentCultureIgnoreCase")} | sort -Descending LastWriteTime | select -First 1 
			"[OK] : Transcript service file retrieved : {0}" -f $logFilePath
		} catch {
			"[ERROR] : Fail to retrieve transcript service backup file : {0}" -f $_.Exception[0].Message
			Stop-Transcript
			exit
		}
	}
	"Components backup" {
		if($Debug) { "Retrieve Transcript_customComponents last file based on last write time" }
		try {
			$logFilePath = Get-ChildItem $DSCExport -ErrorAction Stop | ?{$_.name.StartsWith("Transcript_customComponents","CurrentCultureIgnoreCase")} | sort -Descending LastWriteTime | select -First 1 
			"[OK] : Transcript components file retrieved : {0}" -f $logFilePath
		} catch {
			"[ERROR] : Fail to retrieve transcript components backup file : {0}" -f $_.Exception[0].Message
			Stop-Transcript
			exit
		}
	}
}

#Read the log file content
try {
	$logContent = Get-Content -Path "$DSCExport\$logFilePath" -Encoding UTF8 -ErrorAction Stop
	"[OK] : Transcript file content retrieved"
} catch {
	"[ERROR] : Fail to retrieve transcript file content : {0}" -f $_.Exception[0].Message
	Stop-Transcript
	exit
}

#Parse the log file to retrieve backup duration
try {
	$backupStartDate = ($logContent[2] -split(":"))[1] -replace(" ","")
	$backupStart = [datetime]::ParseExact($backupStartDate, "yyyyMMddHHmmss", $null)
	if($Debug) { "backupStart: {0}" -f $backupStart }
} catch {
	"[ERROR] : Fail to retrieve backup start date from file {0}\{1} : {2}" -f $DSCExport,$logFilePath,$_.Exception[0].Message
	Stop-Transcript
	exit
}

try {
	$backupEndDate = ($logContent[-2] -split(":"))[1] -replace(" ","")
	$backupEnd = [datetime]::ParseExact($backupEndDate, "yyyyMMddHHmmss", $null)
	if($Debug) { "backupEnd: {0}" -f $backupEnd }
} catch {
	"[ERROR] : Fail to retrieve backup end date from file {0}\{1} : {2}" -f $DSCExport,$logFilePath,$_.Exception[0].Message
	Stop-Transcript
	exit
}
try {
	$backupDuration = New-TimeSpan -Start $backupStart -End $backupEnd -ErrorAction Stop
	if($Debug) { "backupDuration: {0}" -f $backupDuration }
} catch {
	"[ERROR] : Fail to retrieve backup duration : {0}" -f $_.Exception[0].Message
	Stop-Transcript
	exit
}
#Parse the log file to make sure script ended successfully
if(!([string]::IsNullOrEmpty($logContent))) {
	if(($logContent[-1]).endswith("*********") -eq $true) {
		$scriptException = $false 
	} else {
		$scriptException = $true
	}
	if($Debug) { "scriptException: $scriptException" }
} else {
	"[ERROR] : Fail to retrieve data from transcript file"
	Stop-Transcript
	exit
}

#Parse the log file and analyze it line by line
foreach ($line in $logContent) {
	if($Debug) { "Processing line: $line" }

    #Check for new module update availability
    if ($line -match "There is a newer version of the 'Microsoft365DSC' module available on the gallery.") {
        $newModuleUpdateAvailable = $true
		if($Debug) { '$line -match "There is a newer version of the "Microsoft365DSC" module available on the gallery."' }
    }
	
	#Check for components list corresponding to backup planned
    if ($line -match "(?:\[INFO\] : Exporting Microsoft 365 configuration for Components:).*(?=)") {
        $componentsPlannedToBackup = ($matches[0] -split(":"))[2] -replace (" ","") -split(",")
		if($Debug) { '$line -match "(?:\[INFO\] : Exporting Microsoft 365 configuration for Components:).*(?=)' }
    }
	
    #Check for component export start
	if ($line -match "Exporting Microsoft 365 configuration for Components: (\w+)") {
		$currentComponent = $matches[1]
		if($Debug) { '$line -match "Exporting Microsoft 365 configuration for Components: (\w+)"' }	
	} 

    #Check for success 
    if ($line -match "✅" -and $currentComponent) {
		if($Debug) {'$line -match "✅" -and $currentComponent' }
        if ($currentComponent -notin $successfulComponents) {
			if($Debug) { "$currentComponent -notin successfulComponents" }
            $successfulComponents += $currentComponent
			if($Debug) { "$currentComponent added to successfulComponents" }
        }
    }

    #Check for errors (Erreur de terminaison or other error indicators)
    if ($line -match "Erreur de terminaison|Error|Failed|Exception|Unable|Cannot|accessDenied|❌|🟡" -and (!($line -match ("No application to sign out from|Get-Failed|ErrorActionPreference")))) {
		if($Debug) { '$line -match "Erreur de terminaison|Error|Failed|Exception|Unable|Cannot|accessDenied|❌|🟡" -and (!($line -match ("No application to sign out from|Get-Failed|ErrorActionPreference")))' }
        if ($currentComponent -and $currentComponent -notin $failedComponents) {
			if($Debug) { "$currentComponent exists and -notin failedComponents" }
            $failedComponents += $currentComponent
			if($Debug) { "failedComponents : {0}" -f $($failedComponents -join(",")) }
        }
    }
}

#Remove components that are in both successful and failed lists (if any)
$successfulComponents = $successfulComponents | Where-Object {$_ -notin $failedComponents }
if($Debug) { "successfulComponents : {0}" -f $($successfulComponents -join(",")) }

#Components not backed-up due to terminating error or process stop
$interruptedComponents = $componentsPlannedToBackup | Where-Object {$_ -notin $successfulComponents -and $_ -notin $failedComponents }
if($Debug) { "interruptedComponents : {0}" -f $($interruptedComponents -join(",")) }

#Calculate global success and failure percentages
$totalComponents = $successfulComponents.Count + $failedComponents.Count + $interruptedComponents.Count
if($Debug) { "totalComponents : {0}" -f $totalComponents }
if ($totalComponents -gt 0) {
    $successPercentage = ($successfulComponents.Count / $totalComponents) * 100
	if($Debug) { "successPercentage : {0}" -f $successPercentage }
    $failurePercentage = ($failedComponents.Count / $totalComponents) * 100
	if($Debug) { "failurePercentage : {0}" -f $failurePercentage }
	$interruptedPercentage = ($interruptedComponents.Count / $totalComponents) * 100
	if($Debug) { "interruptedPercentage : {0}" -f $interruptedPercentage }
} else {
    $successPercentage = 0
    $failurePercentage = 0
	$interruptedPercentage = 0
	if($Debug) { "totalComponents : 0 - percentages : 0" }
}

[int]$GlobalSuccessPercentage = "$([math]::Round($successPercentage, 2))"
if($Debug) { "GlobalSuccessPercentage : {0}" -f $GlobalSuccessPercentage }
[int]$GlobalFailurePercentage = "$([math]::Round($failurePercentage, 2))"
if($Debug) { "GlobalFailurePercentage : {0}" -f $GlobalFailurePercentage }
[int]$GlobalInterruptedPercentage = "$([math]::Round($interruptedPercentage, 2))"
if($Debug) { "GlobalInterruptedPercentage : {0}" -f $GlobalInterruptedPercentage }

if($scriptException -eq $false) {
	if($GlobalSuccessPercentage -ge 75) {
		$globalSuccess = "SUCCESS"
	} elseif(($GlobalSuccessPercentage -lt 75) -and ($GlobalSuccessPercentage -gt 50)) {
		$globalSuccess = "WARNING"
	} elseif($GlobalSuccessPercentage -le 50) {
		$globalSuccess = "FAILURE"
	}
} else {
	$globalSuccess = "INTERRUPTED"
}

if($Debug) { "globalSuccess : {0}" -f $globalSuccess }

#Generate the report
if($BackupType -eq "Full backup" -or $BackupType -eq "Components backup") {
	$Title = "$BackupType Report"
} elseif($BackupType -eq "Services backup") {
	$Title = "$Service Backup Report"
}

if($Debug) { "Generate HTML report" }
$Report = @"
$Title
-------------

Global Backup Job Status: $globalSuccess

Successfull backup :
$($successfulComponents -join "`n")

Failed backup :
$($failedComponents -join "`n")

Interrupted backup : 
$($interruptedComponents -join "`n")

Global Success Percentage: $($GlobalSuccessPercentage)%
Global Failure Percentage: $($GlobalFailurePercentage)%
Global Interrupted Percentage: $($GlobalInterruptedPercentage)%

Backup start : $(Get-Date $backupStart -format "dd/MM/yyyy HH:mm:ss")
Backup end : $(Get-Date $backupEnd -format "dd/MM/yyyy HH:mm:ss")
Backup duration : $($backupDuration.days) days, $($backupDuration.hours) hours, $($backupDuration.minutes) minutes

New PowerShell Module Update Available: $(if($newModuleUpdateAvailable) { "Yes" } else { "No" })
"@		

#Output the report
""
Write-Output $Report
""

#Export txt report to a file
switch($BackupType) {
	"Full backup" { 
		try {
			$Report | Out-File -FilePath "$DSCExport\Backup_FullBackup_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt" -Encoding UTF8 -ErrorAction Stop 
			"[OK] : Backup_FullBackup_Report txt file exported to : {0}" -f $DSCExport
		} catch {
			"[ERROR] : Fail to generate full backup report txt file : {0}" -f $_.Exception[0].Message
		}
	}
	"Services backup" { 
		try {
			$Report | Out-File -FilePath "$DSCExport\Backup_$Service`_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt" -Encoding UTF8 -ErrorAction Stop 
			"[OK] : Backup_$Service`_Report txt file exported to : {0}" -f $DSCExport
		} catch {
			"[ERROR] : Fail to generate services backup report txt file : {0}" -f $_.Exception[0].Message
		}
	}
	"Components backup" { 
		try {
			$Report | Out-File -FilePath "$DSCExport\Backup_customComponents_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt" -Encoding UTF8 -ErrorAction Stop 
			"[OK] : Backup_customComponents_Report txt file exported to : {0}" -f $DSCExport
		} catch {
			"[ERROR] : Fail to generate components backup report txt file : {0}" -f $_.Exception[0].Message
		}
	}
}

Stop-Transcript