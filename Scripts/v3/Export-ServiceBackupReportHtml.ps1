param (
	[Parameter(Mandatory=$false)]
	[ValidateSet("Custom","Meteo","Heartbeat","Smiley","Business")]
	[string]$EmojiType,
	[Parameter(Mandatory=$false)]
	[switch]$DebugEnabled
)

if(!($EmojiType)) {
	$EmojiType = "Custom"
}

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
$ServicesName = @()
$ServicesReport = @()
$LastServiceTxtBackupFileNull = $false
$Emoji = ""
$RepositoryPath = "https://raw.githubusercontent.com/boupipou/public-boupipou-rip-o/refs/heads/main"
$RepositoryPathPictures = "$RepositoryPath/Pictures"

#Script location retrieved automatically
$ScriptLocation = $PSScriptRoot
#If no script location can be retrieved automatically, script location will be current location
if(!($PSScriptRoot -match "\\")) { $ScriptLocation = Get-Location }

#Script name retrieved automatically
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0]
#If no script name can be retrieved automatically, script name will be defined as below
if(!($MyInvocation.MyCommand.Name -match "\w")) { $ScriptName = "Export-ServiceStatusReportHtml" } 

$M365DSCInformations = "Get-M365DSCInformations.json"

#Configuration files directory
$DSCExport = $ScriptLocation

#Email and connection settings
if(Test-Path "$DSCExport\$M365DSCInformations") {
	$InformationsContent = Get-Content "$DSCExport\$M365DSCInformations" | ConvertFrom-Json
	if(!([string]::IsNullOrEmpty($InformationsContent))) {
		$Customer = $InformationsContent | select -ExpandProperty Customer
		$TenantId = $InformationsContent | select -ExpandProperty TenantId
		$TenantName = $InformationsContent | select -ExpandProperty TenantName
		$AppIdExport = $InformationsContent | select -ExpandProperty AppIdExport
		$AppIdImport = $InformationsContent | select -ExpandProperty AppIdImport
		$certSubject = $InformationsContent | select -ExpandProperty certSubject
		$Sender = $InformationsContent | select -ExpandProperty Sender
		if($Debug) {
			$Recipient = $InformationsContent | select -ExpandProperty Cc
			$Recipient2 = $InformationsContent | select -ExpandProperty Cc
			$Cc = $InformationsContent | select -ExpandProperty Cc
		} else {
			$Recipient = $InformationsContent | select -ExpandProperty Recipient
			$Recipient2 = $InformationsContent | select -ExpandProperty Recipient2
			$Cc = $InformationsContent | select -ExpandProperty Cc
		}

	} else {
		"[ERROR] : Fail to retrieve file {0}\{1} content" -f $DSCExport,$M365DSCInformations
		Stop-Transcript
		exit
	}
} else {
	"[ERROR] : File {0}\{1} does not exist" -f $DSCExport,$M365DSCInformations
	Stop-Transcript
	exit
}

$FormattedDateForDirectoryFormat = Get-Date -Format "yyyyMMdd"
$FormattedDate = $(Get-Date).ToString("dd_MM_yyyy_HH-mm-ss")
$dateTime = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"

#Email settings
if($Debug) {
	$Recipient = $Cc
	$Recipient2 = $Cc
	$Cc = $Cc
} else {
	$Recipient = $Recipient
	$Recipient2 = $Recipient2
	$Cc = $Cc
}
#Connection
#Retrieve certificate stored under local machine certificate store
$cert = Get-ChildItem Cert:\LocalMachine\My\ | ?{$_.Subject.StartsWith("CN=$certSubject")}

#Transcript file
$TranscriptFile = "$DSCExport\Transcript_$ScriptName`_$FormattedDate.log"

Start-Transcript -Path $TranscriptFile -Force

#Services' names used for backup
$AzureADServiceName = "AzureAD"
$ExchangeOnlineServiceName = "Exchange"
$IntuneServiceName = "Intune"
$SharePointServiceName = "SharePoint"
$TeamsServiceName = "Teams"
$OneDriveServiceName = "OneDrive"
$SecurityComplianceServiceName = "SecurityCompliance"
if($Debug) { "Services : {0},{1},{2},{3},{4},{5},{6}" -f $AzureADServiceName,$ExchangeOnlineServiceName,$IntuneServiceName,$SharePointServiceName,$TeamsServiceName,$OneDriveServiceName,$SecurityComplianceServiceName }
$ServicesName = @($AzureADServiceName,$ExchangeOnlineServiceName,$IntuneServiceName,$SharePointServiceName,$TeamsServiceName,$OneDriveServiceName,$SecurityComplianceServiceName)
if($Debug) { "ServicesName : {0}" -f $ServicesName }

if($Debug) { 'Retrieve each Backup_$Service last txt file based on last write time' }	

if($Debug) { "Generate HTML report" }

if($Custom) {
	
$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Services Status Report</title>
    <style>
        body { font-family: ui-sans-serif, -apple-system, system-ui, Segoe UI, Helvetica, Apple Color Emoji, Arial, sans-serif, Segoe UI Emoji, Segoe UI Symbol; font-size: 14px; margin: 20px; padding: 20px; color: #333; }
        h2 { color: #333333; margin-top: 20px; }
		h3 { color: #333333; margin-top: 20px; }
        table { text-align: center; width: 50%; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.2);}
        td { padding: 12px; border-bottom: 1px solid #ddd; }
        th { background-color: #f0f0f0; padding: 12px; text-align: center; }
		.tdemoji { font-size: 30px; vertical-align: middle; line-height: 2; }
		.tdimg { width: 100px; height: 100px; vertical-align: middle; line-height: 2; }
		.state { font-size: 15px; }
		.success { background-color: green; color: white; }
		.warning { background-color: orange; color: white; }
		.interrupted { background-color: black; color: white; }
		.failed { background-color: red; color: white; }
		.notfound { background-color: grey; color: white; }
	</style>
</head>
<body>
    <h2>Microsoft 365 Services Status Report</h2>
    <table>
        <tr>
			<th></th>
            <th>Service</th>
            <th>Status</th>
            <th>Success</th>
			<th>Failure</th>
			<th>Interrupted</th>
			<th>Start</th>
			<th>End</th>
        </tr>
"@
	
} else {
	
$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft 365 Services Status Report</title>
    <style>
        body { font-family: ui-sans-serif, -apple-system, system-ui, Segoe UI, Helvetica, Apple Color Emoji, Arial, sans-serif, Segoe UI Emoji, Segoe UI Symbol; font-size: 14px; margin: 20px; padding: 20px; color: #333; }
        h2 { color: #333333; margin-top: 20px; }
		h3 { color: #333333; margin-top: 20px; }
        table { text-align: center; width: 30%; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.2); }
        td { padding: 12px; border-bottom: 1px solid #ddd; }
        th { background-color: #f0f0f0; padding: 12px; text-align: center; }
		.emoji { font-size: 30px; line-height: 2; display: flex; justify-content: center; align-items: center; }
	</style>
</head>
<body>
    <h2>Microsoft 365 Services Status Report</h2>
    <table>
        <tr>
			<th></th>
            <th>Service</th>
            <th>Status</th>
            <th>Success</th>
			<th>Failure</th>
			<th>Interrupted</th>
			<th>Start</th>
			<th>End</th>
        </tr>
"@

}

Foreach($ServiceName in $ServicesName) {
	try {
		$logFilePath = Get-ChildItem $DSCExport -ErrorAction Stop | ?{$_.name.StartsWith("Backup_$ServiceName") -and $_.name.EndsWith(".txt")} | sort -Descending LastWriteTime | select -First 1 
		if(!([string]::IsNullOrEmpty($logFilePath))) {
			"[OK] : Last service {0} backup txt file retrieved : {1}" -f $ServiceName,$logFilePath
			$LastServiceTxtBackupFileNull = $false
		} else {
			"[INFO] : Fail to retrieve last service backup txt file : {0}" -f $ServiceName
			$LastServiceTxtBackupFileNull = $true
			$globalSuccess = "NOTFOUND"
		}
	} catch {
		"[ERROR] : Fail to retrieve last service {0} backup txt file : {1}" -f $ServiceName,$_.Exception[0].Message	
		$LastServiceTxtBackupFileNull = $true
		$globalSuccess = "NOTFOUND"
	}
	if($Debug) { "LastServiceTxtBackupFileNull : {0}" -f $LastServiceTxtBackupFileNull }
	if($Debug) { "Target backup txt file for {0} : {1}\{2}" -f $ServiceName,$DSCExport,$logFilePath }

	if($LastServiceTxtBackupFileNull -ne $true) {
		#Read the log file content
		if(!([string]::IsNullOrEmpty($logFilePath))) {
			try {
				$logContent = Get-Content -Path "$DSCExport\$logFilePath" -Encoding Default	
				"[OK] : Content retrieved from file {0}\{1}" -f $DSCExport,$logFilePath
			} catch {
				"[ERROR] : Fail to retrieve content from file {0}\{1} : {2}" -f $DSCExport,$logFilePath,$_.Exception[0].Message
				$LastServiceTxtBackupFileNull = $true
				$globalSuccess = "NOTFOUND"
			}
		} else {
			"[INFO] : Fail to retrieve data from backup txt file : {0}" -f $logFilePath
		}

		#Parse the file content
		if(!([string]::IsNullOrEmpty($logContent))) {
			$backupStart = (($logContent | Select-String "backup start") -split(": "))[1]
			if($Debug) { "backupStart: {0}" -f $backupStart }
			$backupEnd = (($logContent | Select-String "backup end") -split(": "))[1]
			if($Debug) { "backupEnd: {0}" -f $backupEnd }

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
		} else {
			$LastServiceTxtBackupFileNull = $true
			$globalSuccess = "NOTFOUND"
		}
	}
	
	if($Debug) { "globalSuccess: {0}" -f $globalSuccess }
	
	if(!$Custom) { #If custom image is selected, change the job statuses text to another color
		switch -regex ($globalSuccess) {
			"SUCCESS" {
				$Class = '<span style="color: green;">'
			}
			"FAILURE" {
				$Class = '<span style="color: red;">'
			}
			"NOTFOUND" {
				$Class = '<span style="color: gray;">'
			}
			"WARNING" {
				$Class = '<span style="color: orange;">'
			}
			"INTERRUPTED" {
				$Class = '<span style="color: black;">'
			}
		}
		if($Debug) { "class: {0}" -f $Class }
	}
	
	switch ($EmojiType) {
		"Custom" {
			$Emoji = @"
			<img class="tdimg" src="$RepositoryPathPictures/$($globalSuccess.ToLower())`.png"></img>
"@
		}
		"Meteo" {
			if($globalSuccess -eq "SUCCESS") {
				$Emoji = "🌞"
			} elseif($globalSuccess -eq "FAILURE") {
				$Emoji = "⛈️"
			} elseif($globalSuccess -eq "WARNING") {
				$Emoji = "⛅"
			} elseif($globalSuccess -eq "INTERRUPTED") {
				$Emoji = "🌪️"
			} elseif($globalSuccess -eq "NOTFOUND") {
				$Emoji = "☁️"
			}
		}
		"Hearbeat" {
			if($globalSuccess -eq "SUCCESS") {
				$Emoji = "💖"
			} elseif($globalSuccess -eq "FAILURE") {
				$Emoji = "🖤"
			} elseif($globalSuccess -eq "WARNING") {
				$Emoji = "💛"
			} elseif($globalSuccess -eq "INTERRUPTED") {
				$Emoji = "💔"
			} elseif($globalSuccess -eq "NOTFOUND") {
				$Emoji = "🤍"
			}
		}
		"Smiley" {
			if($globalSuccess -eq "SUCCESS") {
				$Emoji = "😜"
			} elseif($globalSuccess -eq "FAILURE") {
				$Emoji = "😭"
			} elseif($globalSuccess -eq "WARNING") {
				$Emoji = "😰"
			} elseif($globalSuccess -eq "INTERRUPTED") {
				$Emoji = "🥴"
			} elseif($globalSuccess -eq "NOTFOUND") {
				$Emoji = "👻"
			}
		}
		"Business" {
			if($globalSuccess -eq "SUCCESS") {
				$Emoji = "✅"
			} elseif($globalSuccess -eq "FAILURE") {
				$Emoji = "❌"
			} elseif($globalSuccess -eq "WARNING") {
				$Emoji = "⚠️"
			} elseif($globalSuccess -eq "INTERRUPTED") {
				$Emoji = "🚧"
			} elseif($globalSuccess -eq "NOTFOUND") {
				$Emoji = "⛔"
			}
		}
	}	
		
	#Define the data
	if($LastServiceTxtBackupFileNull -ne $true) {
		$ServiceReport = @([PSCustomObject]@{Emoji = $Emoji; Service = $ServiceName; Status = $globalSuccess; "Success Percentage" = $successPercentage; "Failure Percentage" = $failurePercentage; "Interrupted Percentage" = $interruptedPercentage; Start = $backupStart; End = $backupEnd})
	} else {
		$ServiceReport = @([PSCustomObject]@{Emoji = $Emoji; Service = $ServiceName; Status = $globalSuccess; "Success Percentage" = "-%"; "Failure Percentage" = "-%"; "Interrupted Percentage" = "-%"; Start = "NO DATA"; End = "NO DATA"})
	}
	
	if($Custom) {	
	$class = $globalSuccess.ToLower()
	    $htmlContent += @"
			<td><span class="tdemoji">$($ServiceReport.Emoji)</span></td>
            <td>$($ServiceReport.Service)</td>
			<td><span class="$class"><b class="state">$($ServiceReport.Status)</b></span></td>
			<td>$($ServiceReport."Success Percentage")</td>
			<td>$($ServiceReport."Failure Percentage")</td>
			<td>$($ServiceReport."Interrupted Percentage")</td>
			<td>$($ServiceReport.Start)</td>
			<td>$($ServiceReport.End)</td>
        </tr>
"@
	} else {
		$htmlContent += @"
        <tr>
			<td><span class="emoji">$($ServiceReport.Emoji)</span></td>
            <td>$($ServiceReport.Service)</td>
			<td>$Class<b>$($ServiceReport.Status)</b></span></td>
			<td>$($ServiceReport."Success Percentage")</td>
			<td>$($ServiceReport."Failure Percentage")</td>
			<td>$($ServiceReport."Interrupted Percentage")</td>
			<td>$($ServiceReport.Start)</td>
			<td>$($ServiceReport.End)</td>
        </tr>
"@
	}
}


$htmlContent += @"
    </table>
</body>
</html>
"@

#Save the HTML content to a file
$htmlReportFilePath = "$DSCExport\$ScriptName`_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

try {
	$htmlContent | Out-File -FilePath $htmlReportFilePath -Encoding UTF8 -Force -ErrorAction Stop
	"[OK] : HTML report generated successfully : {0}" -f $htmlReportFilePath
} catch {
	"[ERROR] : Fail to generate html report {0} : {1}" -f $htmlReportFilePath,$_.Exception[0].Message
}

if(!([string]::IsNullOrEmpty($htmlReportFilePath))) {
	try {
		$BodyContent = Get-Content $htmlReportFilePath -Encoding UTF8 -Raw -ErrorAction Stop
		"[OK] : Html content retrieved from file {0}" -f $htmlReportFilePath
	} catch {
		"[ERROR] : Fail to retrieve content from file {0} : {1}" -f $htmlReportFilePath,$_.Exception[0].Message
	}
} else {
	"[ERROR] : Html file is null or empty"
}

#Connect to Graph to send email
try {
	Connect-MgGraph -ClientId $AppIdExport -Certificate $cert -TenantId $TenantName -ErrorAction Stop | Out-Null
	"[OK] : Connected to Microsoft Graph"
} catch {
	"[ERROR] : Fail to run cmdlet 'Connect-MgGraph' : {0}" -f $_.Exception[0].Message
	break
}
	
$Subject = "$Customer - M365DSC Backup services report"
$SendMailParams = @{
	Message = @{
		Subject = $Subject
		Body = @{
			ContentType = "html"
			Content = $BodyContent
		}
		ToRecipients = @(
			@{
				EmailAddress = @{
					Address = $Recipient
				}
			}			
			@{
				EmailAddress = @{
					Address = $Recipient2
				}
			}
		)
		CcRecipients = @(
			@{
				EmailAddress = @{
					Address = $Cc
				}
			}
		)
	}

SaveToSentItems = "false"
}
"[INFO]: Sending email From:{0} - To:{1},{2} - Cc:{3} - Subject:{4}" -f $Sender,$Recipient,$Recipient2,$Cc,$Subject

try {
	Send-MgUserMail -UserId $Sender -BodyParameter $SendMailParams -ErrorAction Stop
	"[OK] : Email sent From:{0} - To:{1},{2} - Cc:{3} - Subject:{4}" -f $Sender,$Recipient,$Recipient2,$Cc,$Subject
} catch {
	"[ERROR]: Report could not be sent by mail : {0}" -f $_.Exception[0].Message
}

Disconnect-MgGraph | Out-Null

Stop-Transcript