[CmdletBinding()]
param(
	[Parameter(Mandatory=$false)]
	[ValidateSet("AzureAD","Exchange","Intune","Teams","SharePoint","OneDrive","SecurityCompliance")]
	[string]$ServiceRequested,
	[Parameter(Mandatory=$false)]
	[array]$ComponentRequested,
	[Parameter(Mandatory=$false)]
	[switch]$FullBackup=$false,
	[Parameter(Mandatory=$false)]
	[switch]$Help=$false,
	[switch]$DebugEnabled,
	[switch]$ManualBackup=$false
)

<#Parameters :
-Start script to backup all services with all their components in one file : specify -FullBackup
If full backup is requested, no specific service or component can be specified
-Start script to backup one service with all their components : specify -ServiceRequested Service1
-Start script to backup one or several components of one or several service(s) : specify -ComponentRequested Component1,Component2
-Help : provides guidance about script execution and parameters
-Debug : provides additional debug lines not displayed in production mode
#>

if($DebugEnabled) {
	$Debug = $true
} else {
	$Debug = $false
}

$Services = @()
$AADComponents = @()
$EXOComponents = @()
$INTUNEComponents = @()
$SPOComponents = @()
$TEAMSComponents = @()
$ODComponents = @()
$SCComponents = @()
$AllComponents = @()
$Components = $null
$ComponentsToBackup = $null
$BackedUpServicesAndComponents = @()
$DoNotRemoveExceptions = @()
$ComponentsInput = @()

#Script location retrieved automatically
$ScriptLocation = $PSScriptRoot
#If no script location can be retrieved automatically, script location will be current location
if(!($PSScriptRoot -match "\\")) { $ScriptLocation = Get-Location }

#Script name retrieved automatically
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0]
#If no script name can be retrieved automatically, script name will be defined as below
if(!($MyInvocation.MyCommand.Name -match "\w")) { $ScriptName = "Export-DSCConfiguration" } 

#Reporting dependencies scripts and files
$M365DSCGraphicalUIForBackup = "Create-ServiceBackupForm.ps1"
$M365DSCGraphicalUIForRestoration = "Create-ServiceRestorationForm.ps1"
$M365DSCBackupScript = "Export-DSCConfiguration.ps1"
$M365DSCHtmlReportScriptByService = "Export-ServiceBackupReportHtml.ps1"
$M365DSCTxtReportScriptByService = "Export-ServiceBackupReportTxt.ps1"
$M365DSCHtmlReportGlobal = "Export-ServiceStatusReportHtml.ps1"	
$M365DSCRestorationScript = "Import-DSCConfiguration.ps1"
$M365DSCSendMail = "Send-BackupJobReportsByMail.ps1"
$M365DSCUpdateComponentsList = "Update-M365DSCComponentsList.ps1"
$M365DSCUpdateModules = "Update-M365DSCModuleAndDependencies.ps1"
$M365DSCComponentsList = "Get-M365DSCComponentsList_v1.txt"
$M365DSCInformations = "Get-M365DSCInformations.json"
$M365DSCCertificatePfx = "M365DSC.pfx"
$M365DSCCertificateCer = "M365DSC.cer"
$M365DSCIco = "dsc.ico"

#Email and connection settings
if(Test-Path "$ScriptLocation\$M365DSCInformations") {
	$InformationsContent = Get-Content "$ScriptLocation\$M365DSCInformations" | ConvertFrom-Json
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
		"[ERROR] : Fail to retrieve file {0}\{1} content" -f $ScriptLocation,$M365DSCInformations
		Stop-Transcript
		exit
	}
} else {
	"[ERROR] : File {0}\{1} does not exist" -f $ScriptLocation,$M365DSCInformations
	Stop-Transcript
	exit
}

#Retrieve certificate stored under local machine certificate store
$cert = Get-ChildItem Cert:\LocalMachine\My\ | ?{$_.Subject.StartsWith("CN=$certSubject")}
$DoNotRemoveExceptions = @(
$M365DSCGraphicalUIForBackup,
$M365DSCGraphicalUIForRestoration,
$M365DSCBackupScript,
$M365DSCHtmlReportScriptByService,
$M365DSCTxtReportScriptByService,
$M365DSCHtmlReportGlobal,
$M365DSCRestorationScript,
$M365DSCSendMail,
$M365DSCUpdateComponentsList,
$M365DSCUpdateModules,
$M365DSCInformations,
$M365DSCComponentsList,
$M365DSCCertificatePfx,
$M365DSCCertificateCer,
$M365DSCIco
)

if($Debug) { "Exception list of files not removed: {0}" -f $($DoNotRemoveExceptions -join(",")) }

Set-Location $ScriptLocation

$FormattedDateForDirectoryFormat = Get-Date -Format "yyyyMMdd"
$FormattedDate = $(Get-Date).ToString("dd_MM_yyyy_HH-mm-ss")

#Transcript script execution file
$TranscriptGlobalFile = "$ScriptLocation\Transcript_$ScriptName`_$FormattedDate.log"
	
Start-Transcript -Path $TranscriptGlobalFile -Force

#If no component is specified, set services names and components to backup
if([string]::IsNullOrEmpty($ComponentRequested)) {
	if($Debug) { "No component specified" }

	#Services' names used for backup
	$AzureADServiceName = "AzureAD"
	$ExchangeOnlineServiceName = "Exchange"
	$IntuneServiceName = "Intune"
	$SharePointServiceName = "SharePoint"
	$TeamsServiceName = "Teams"
	$OneDriveServiceName = "OneDrive"
	$SecurityComplianceServiceName = "SecurityCompliance"
	if($Debug) { "Services : {0},{1},{2},{3},{4},{5},{6}" -f $AzureADServiceName,$ExchangeOnlineServiceName,$IntuneServiceName,$SharePointServiceName,$TeamsServiceName,$OneDriveServiceName,$SecurityComplianceServiceName }

	#File containing all components backed-up
	$ComponentsFileName = "Get-M365DSCComponentsList_v1.txt"
	$ComponentsFile = "$ScriptLocation\$ComponentsFileName"
	if($Debug) { "Components list file : {0}" -f $ComponentsFile }

	try {
		$Components = Get-Content $ComponentsFile -ErrorAction Stop | select -Unique
		"[INFO] : {0} Components retrieved from path {1}" -f $($Components | Measure-Object | select -ExpandProperty count),$ComponentsFile
	} catch {
		"[ERROR] : Fail to retrieve components list from path {0} : {1}" -f $ComponentsFile,$_.Exception[0].Message
		Stop-Transcript
		exit
	}

	#Retrieve each component for each service based on prefix
	$AADComponents += $Components | ?{$_.startswith("AAD","CurrentCultureIgnoreCase")} | sort
	if($Debug) { "AADComponents : {0}" -f $($AADComponents -join(",")) }
	$EXOComponents += $Components | ?{$_.startswith("EXO","CurrentCultureIgnoreCase")} | sort
	if($Debug) { "EXOComponents : {0}" -f $($EXOComponents -join(",")) }
	$INTUNEComponents += $Components | ?{$_.startswith("INTUNE", "CurrentCultureIgnoreCase")} | sort
	if($Debug) { "INTUNEComponents : {0}" -f $($INTUNEComponents -join(",")) }
	$SPOComponents += $Components | ?{$_.startswith("SPO","CurrentCultureIgnoreCase")} | sort
	if($Debug) { "SPOComponents : {0}" -f $($SPOComponents -join(",")) }
	$TEAMSComponents += $Components | ?{$_.startswith("TEAMS","CurrentCultureIgnoreCase")} | sort
	if($Debug) { "TEAMSComponents : {0}" -f $($TEAMSComponents -join(",")) }
	$ODComponents += $Components | ?{$_.startswith("OD","CurrentCultureIgnoreCase")} | sort
	if($Debug) { "ODComponents : {0}" -f $($ODComponents -join(",")) }
	$SCComponents += $Components | ?{$_.startswith("SC","CurrentCultureIgnoreCase")} | sort
	if($Debug) { "SCComponents : {0}" -f $($SCComponents -join(",")) }
	$AllComponents = $Components | select -Unique | sort
	if($Debug) { "AllComponents : {0}" -f $($AllComponents -join(",")) }
}

#Display help about script execution
switch($Help) {
	$true {
		Clear-Host
		$Timer = 15
		
		Write-Host "[INFO]: This script can run three parameters that" -ForeGroundColor Yellow -NoNewLine;Write-Host " cannot " -ForeGroundColor Red -NoNewLine;Write-Host "be used at the same time" -ForeGroundColor Yellow
		Write-Host "[INFO]: To perform a full backup : .\PathToScript\Export-DSCConfiguration.ps1" -ForeGroundColor Gray -NoNewLine;Write-Host " -FullBackup" -ForeGroundColor Yellow
		Write-Host "[INFO]: To back up specific service and all of their components : .\PathToScript\Export-DSCConfiguration.ps1" -ForeGroundColor Gray -NoNewLine;Write-Host " -ServiceRequested Service" -ForeGroundColor Yellow
		Write-Host "[INFO]: To back up specific components : .\PathToScript\Export-DSCConfiguration.ps1" -ForeGroundColor Gray -NoNewLine;Write-Host " -ComponentRequested Component1,Component2" -ForeGroundColor Yellow
		Write-Host "[INFO]: Services backup currently supported : " -ForeGroundColor Gray -NoNewLine;Write-Host "AzureAD,Exchange,Intune,SharePoint,Teams,OneDrive,SecurityCompliance" -ForeGroundColor Yellow
		""
		""
		Write-Host "[INFO]: Exhaustive listing of services and components will open in $Timer seconds in a web browser" -ForeGroundColor Gray
		Write-Host "[INFO]: Backup files will be found as .ps1 files that need to be executed as scripts in order to provide the restauration file" -ForeGroundColor Yellow
		Write-Host "[INFO]: Transcript files will be generated and contain information about : " -ForeGroundColor Gray
		Write-Host "`to Start time " -ForeGroundColor Gray
		Write-Host "`to End time " -ForeGroundColor Gray
		Write-Host "`too Exported Services " -ForeGroundColor Gray
		Write-Host "`too Exported Components " -ForeGroundColor Gray
		Write-Host "`too Components export duration" -ForeGroundColor Gray
		Write-Host "`tooo PowerShell modules update" -ForeGroundColor Gray
		Write-Host "`tooo Backup success" -ForeGroundColor Gray
		Write-Host "`tooo Encountered errors" -ForeGroundColor Gray
		
		Start-Sleep -Seconds $Timer
		
		Export-M365DSCConfiguration -LaunchWebUI | Out-Null
		""
		
		Stop-Transcript
		exit
	}
	default {
		#Exit the loop and start the script in normal execution
		break
	}
}

#Whether to perform a full backup or split into several services and components
switch($FullBackup) {
	$true { 
		$FullBackup = $true 
		$BackupType = "Full backup"
		Write-Host "[INFO]: Perform a full backup into one file" -ForeGroundColor Yellow
		$ComponentsToBackup = $AllComponents -join (",")
		"[INFO] : Components to backup for full backup job : {0}" -f $ComponentsToBackup
		if($ServiceRequested -or $ComponentRequested) {
			Write-Host "[ERROR]: Full backup requested - you're not allowed to specify any specific service or component" -ForeGroundColor Yellow
			""
			Stop-Transcript
			exit
		}
	}
	default { 
		$FullBackup = $false 
		<#Full backup not requested, every service will be backuped and split into different directories
		Unless a specific service is specified #>
		
		if($ServiceRequested) {
			#If specific service  specified, only this service and their components will be backuped
			$Services = $ServiceRequested
			$BackupType = "Services backup"
			Write-Host "[INFO]: Specific service '$($Services)' requested" -ForeGroundColor Yellow
			#No component should be specified when specifying services
			if($ComponentRequested) {
				Write-Host "[ERROR]: Service backup requested - you're not allowed to specify any specific component" -ForeGroundColor Yellow
				Stop-Transcript
				exit
			}
		} else {
			if(!($ComponentRequested)) {
				#If no specific service is specified, exit
				Write-Host "[INFO]: No specific services requested. Exiting script.'" -ForeGroundColor Yellow
				Stop-Transcript
				exit
			}
		}
		
		if($ComponentRequested) {
			#If specific components are specified, only these components will be backuped
			$ComponentsInput = $ComponentRequested
			$BackupType = "Components backup"
			if($ServiceRequested) {
				Write-Host "[ERROR]: Component backup requested - you're not allowed to specify any specific service" -ForeGroundColor Yellow
				Stop-Transcript
				exit
			}
		}
	}
}

if(!($ComponentRequested)) {
	Foreach($Service in $Services) {
		Switch($Service) {
			$AzureADServiceName { 
				#Service backup
				if($FullBackup -ne $true) { 
					$ComponentsToBackup += $AADComponents -join(",")
				}
			}
			$ExchangeOnlineServiceName { 
				#Service backup
				if($FullBackup -ne $true) { 
					$ComponentsToBackup += $EXOComponents -join(",")
				}
			}
			$IntuneServiceName { 
				#Service backup
				if($FullBackup -ne $true) { 
					$ComponentsToBackup += $INTUNEComponents -join(",")
				}	
			}
			$OneDriveServiceName {
				#Service backup
				if($FullBackup -ne $true) { 
					$ComponentsToBackup += $ODComponents -join(",")
				}
			}
			$SharePointServiceName {
				#Service backup
				if($FullBackup -ne $true) { 
					$ComponentsToBackup += $SPOComponents -join(",")
				}
			}
			$TeamsServiceName {
				#Service backup
				if($FullBackup -ne $true) { 
					$ComponentsToBackup += $TEAMSComponents -join(",")
				}
			}
			$SecurityComplianceServiceName {
				#Service backup
				if($FullBackup -ne $true) {
					$ComponentsToBackup += $SCComponents -join(",")
				}
			}
			default {  
				Write-Host "[ERROR]: No component found for $Service" -ForeGroundColor Yellow
				Stop-Transcript
				exit
			}
		}
	}
}

""
Write-Host "[INFO]: Script starting location '$ScriptLocation'" -ForeGroundColor Yellow

if($FullBackup -eq $true) { 
	#@@@@@@@@@@@@@@@@
	#FULL BACKUP
	#@@@@@@@@@@@@@@@@
	#If full backup is specified, back up every service and component in one file
	
	Write-Host "[INFO]: Starting full backup" -ForeGroundColor Yellow
	
	#Transcript file
	$TranscriptFile = "$ScriptLocation\Transcript_FullBackup_$FormattedDate.log"
	
	#Log the job into transcript file
	Start-Transcript -Path $TranscriptFile -Force
	
	"[INFO] : Exporting Microsoft 365 configuration for Components: {0}" -f $ComponentsToBackup
	
	try {
		Export-M365DSCConfiguration -CertificateThumbprint $cert.Thumbprint -TenantId $TenantName -ApplicationId $AppIdExport -Path "$ScriptLocation\FullBackup_$FormattedDateForDirectoryFormat" -FileName "FullBackup.ps1" -Components $ComponentsToBackup
	} catch {
		"[ERROR] : Fail to perform full backup : {0}" -f $_.Exception[0].Message
	}

	Stop-Transcript
	""
} elseif(!($ComponentRequested)) {
	#@@@@@@@@@@@@@@@@
	#SERVICES BACKUP
	#@@@@@@@@@@@@@@@@
	#If no component is specified, backup by service each component
	
	Write-Host "[INFO]: Starting backup for services " -ForeGroundColor Yellow
	
	Foreach($Service in $Services) {
		#Transcript file
		$TranscriptFile = "$ScriptLocation\Transcript_$Service`_$FormattedDate.log"

		#Log the job into transcript file
		Start-Transcript -Path $TranscriptFile -Force
		
		"[INFO] : Exporting Microsoft 365 configuration for Components: {0}" -f $ComponentsToBackup
		
		Foreach($Component in ($ComponentsToBackup -split(","))) {
			try {
				Export-M365DSCConfiguration -CertificateThumbprint $cert.Thumbprint -TenantId $TenantName -ApplicationId $AppIdExport -Path "$ScriptLocation\$Service`_$FormattedDateForDirectoryFormat" -FileName "$Component.ps1" -Components $Component
			} catch {
				"[ERROR] : Fail to backup component {0} for service {1} : {2}" -f $Component,$Service,$_.Exception[0].Message
			}
		}

		Stop-Transcript
		""
		
		if(Test-Path $TranscriptFile) {
			if($Debug) { "TranscriptFile found {0}" -f $TranscriptFile }
			
			#Send email only if backup is executed manually ; if executed automatically, email report will be handled by another task
			if($ManualBackup) {	
				Foreach($Service in $Services) {
					#Generate txt file used to create html report by service
					try {
						Start-Process powershell.exe -ArgumentList "-File `"$ScriptLocation\$M365DSCTxtReportScriptByService`" -Service `"$Service`" -BackupType `"Services backup`"" -ErrorAction Stop -NoNewWindow #-WindowStyle Hidden
						if($Debug) { 'Start-Process powershell.exe -ArgumentList "-File `"$ScriptLocation\$M365DSCTxtReportScriptByService`" -Service `"$Service`" -BackupType `"Services backup`"" -ErrorAction Stop -NoNewWindow' }
						"[OK] : Script {0} successfully started" -f $M365DSCTxtReportScriptByService
					} catch {
						"[ERROR] : Fail to generate report txt file using script {0} : {1}" -f $M365DSCTxtReportScriptByService,$_.Exception[0].Message
					}
					
					if($Debug) { "Start-Sleep 30s between two script execution" }
					Start-Sleep -Seconds 30
					
					#Generate html file used to report backup service state
					try {
						Start-Process powershell.exe -ArgumentList "-File `"$ScriptLocation\$M365DSCHtmlReportScriptByService`" -Service `"$Service`" -BackupType `"Services backup`"" -ErrorAction Stop -NoNewWindow #-WindowStyle Hidden
						if($Debug) { 'Start-Process powershell.exe -ArgumentList "-File `"$ScriptLocation\$M365DSCHtmlReportScriptByService`" -Service `"$Service`" -BackupType `"Services backup`"" -ErrorAction Stop -NoNewWindow' }
						"[OK] : Script {0} successfully started" -f $M365DSCHtmlReportScriptByService
					} catch {
						"[ERROR] : Fail to generate report html file using script {0} : {1}" -f $M365DSCHtmlReportScriptByService,$_.Exception[0].Message
					}
					
					if($Debug) { "Start-Sleep 30s between two script execution" }
					Start-Sleep -Seconds 30
				
					if($Debug) { "Retrieve last Backup_$Service html file and content based on last write time in order to send through email" }
					
					try {
						$htmlReportFilePath = Get-ChildItem $ScriptLocation -ErrorAction Stop | ?{$_.name.StartsWith("Backup_$Service") -and $_.name.EndsWith(".html")} | sort -Descending LastWriteTime | select -First 1
						"[OK] : Html file retrieved : {0}" -f $htmlReportFilePath
					} catch {
						"[ERROR] : Fail to retrieve html file {0} : {1}" -f $htmlReportFilePath,$_.Exception[0].Message
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
					
					$Subject = "$Customer - $Service backup report"
					$MessageAttachement = [Convert]::ToBase64String([IO.File]::ReadAllBytes($TranscriptFile))
					try {
						$TranscriptContent = Get-Content $TranscriptFile -ErrorAction Stop
						"[OK] : Job transcript file content retrieved from file : {0}" -f $TranscriptFile
					} catch {
						"[ERROR] : Fail to retrieve content from file {0} : {1}" -f $TranscriptFile,$_.Exception[0].Message	
					}
					
					#Error log file contains all services in one file
					#It should be sent after the last service has been completed
					if($Service -eq $Services[-1]) {
						#Check for any error file generated by the script
						$ErrorEntries = @()
						foreach($entry in $TranscriptContent) {
							$ErrorEntries += $entry | ?{$_ -match "file://"} 
						}
						
						if(($null -ne $ErrorEntries) -and ($ErrorEntries)) {
							$ErrorLogFile = (($ErrorEntries -split'{|}') -replace ('file://','') -replace ('\/','\') | group | select -ExpandProperty Name)[1]
							if(Test-Path $ErrorLogFile) {
								if($ErrorLogFile -match "\\") {
									Write-Host "[INFO]: Error file '$ErrorLogFile' found to be added to the logs" -ForeGroundColor Yellow
									$MessageAttachement2 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($ErrorLogFile))
									
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
											Attachments = @(
												@{
													"@odata.type" = "#microsoft.graph.fileAttachment"
													Name = ($TranscriptFile -split"\\")[-1]
													ContentType = "text/plain"
													ContentBytes = $MessageAttachement
												}
												@{
													"@odata.type" = "#microsoft.graph.fileAttachment"
													Name = ($ErrorLogFile -split"\\")[-1]
													ContentType = "text/plain"
													ContentBytes = $MessageAttachement2
												}
											)
										}

									SaveToSentItems = "false"
									}
								}
							} else {
								Write-Host "[ERROR]: Error file '$ErrorLogFile' could not be found" -ForeGroundColor Yellow
							}
						} else {
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
									Attachments = @(
										@{
											"@odata.type" = "#microsoft.graph.fileAttachment"
											Name = ($TranscriptFile -split"\\")[-1]
											ContentType = "text/plain"
											ContentBytes = $MessageAttachement
										}
									)
								}
							SaveToSentItems = "false"	
							}
						}
					} else {
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
								Attachments = @(
									@{
										"@odata.type" = "#microsoft.graph.fileAttachment"
										Name = ($TranscriptFile -split"\\")[-1]
										ContentType = "text/plain"
										ContentBytes = $MessageAttachement
									}
								)
							}
						SaveToSentItems = "false"	
						}
					}

					Write-Host "[INFO]: Sending email From:'$Sender' To:'$Recipient','$Recipient2' Cc:'$Cc' with Subject:'$Subject'" -ForeGroundColor Yellow
				
					#Connect to Graph by the end of the script to avoid any disconnection during process
					try {
						Connect-MgGraph -ClientId $AppIdExport -Certificate $cert -TenantId $TenantName -ErrorAction Stop | Out-Null
						"[OK] : Connected to Microsoft Graph"
					} catch {
						"[ERROR] : Fail to run cmdlet 'Connect-MgGraph' : {0}" -f $_.Exception[0].Message
						break
					}
				
					try {
						Send-MgUserMail -UserId $Sender -BodyParameter $SendMailParams -ErrorAction Stop
						"[OK] : Email sent From:{0} - To:{1},{2} - Cc:{3} - Subject:{4}" -f $Sender,$Recipient,$Recipient2,$Cc,$Subject
					} catch {
						"[ERROR]: Report could not be sent by mail : {0}" -f $_.Exception[0].Message
					}
				} 
			} else {
				if($Debug) { "-ManualBackup not specified, no email or html file to generate" }
			}
		} else {
			"[ERROR] : No transcript file found to generate txt report file"
		}
	}
} else {
	#@@@@@@@@@@@@@@@@@@
	#COMPONENTS BACKUP
	#@@@@@@@@@@@@@@@@@@
	#If components are specified they can come from different services
	
	#Backup each component in a custom directory
	Write-Host "[INFO]: Starting backup for components" -ForeGroundColor Yellow
	
	#Transcript file
	$TranscriptFile = "$ScriptLocation\Transcript_customComponents`_$FormattedDate.log"
	
	Foreach($Component in ($ComponentsInput -split(","))) {
		
		#Log the job into transcript file
		Start-Transcript -Path $TranscriptFile -Force
		
		"[INFO] : Exporting Microsoft 365 configuration for Component: {0}" -f $Component
	
		try {
			Export-M365DSCConfiguration -CertificateThumbprint $cert.Thumbprint -TenantId $TenantName -ApplicationId $AppIdExport -Path "$ScriptLocation\customComponents_$FormattedDateForDirectoryFormat" -FileName "$Component.ps1" -Components $Component
		} catch {
			"[ERROR] : Fail to backup component {0} : {1}" -f $Component,$_.Exception[0].Message
		}
	
		Stop-Transcript
		""
	
		#Generate txt file used to create html report by service
		try {
			Start-Process powershell.exe -ArgumentList "-File `"$ScriptLocation\$M365DSCTxtReportScriptByService`" -BackupType `"Components backup`"" -ErrorAction Stop -NoNewWindow #-WindowStyle Hidden
			if($Debug) { 'Start-Process powershell.exe -ArgumentList "-File `"$ScriptLocation\$M365DSCTxtReportScriptByService`" -BackupType `"Components backup`"" -ErrorAction Stop -NoNewWindow' }
			"[OK] : Script {0} successfully started" -f $M365DSCTxtReportScriptByService
		} catch {
			"[ERROR] : Fail to generate report txt file using script {0} : {1}" -f $M365DSCTxtReportScriptByService,$_.Exception[0].Message
		}
		
		if($Debug) { "Start-Sleep 30s between two script execution" }
		Start-Sleep -Seconds 30
		
		#Generate html file used to report backup service state
		try {
			Start-Process powershell.exe -ArgumentList "-File `"$ScriptLocation\$M365DSCHtmlReportScriptByService`" -BackupType `"Components backup`"" -ErrorAction Stop -NoNewWindow #-WindowStyle Hidden
			if($Debug) { 'Start-Process powershell.exe -ArgumentList "-File `"$ScriptLocation\$M365DSCHtmlReportScriptByService`" -BackupType `"Components backup`"" -ErrorAction Stop -NoNewWindow' }
			"[OK] : Script {0} successfully started" -f $M365DSCHtmlReportScriptByService
		} catch {
			"[ERROR] : Fail to generate report html file using script {0} : {1}" -f $M365DSCHtmlReportScriptByService,$_.Exception[0].Message
		}
		
		if($Debug) { "Start-Sleep 30s between two script execution" }
		Start-Sleep -Seconds 30

		if($Debug) { "Retrieve last Backup_customComponents html file and content based on last write time in order to send through email" }
		
		try {
			$htmlReportFilePath = Get-ChildItem $ScriptLocation -ErrorAction Stop | ?{$_.name.StartsWith("Backup_customComponents") -and $_.name.EndsWith(".html")} | sort -Descending LastWriteTime | select -First 1
			"[OK] : Html file retrieved : {0}" -f $htmlReportFilePath
		} catch {
			"[ERROR] : Fail to retrieve html file {0} : {1}" -f $htmlReportFilePath,$_.Exception[0].Message
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
		
		$Subject = "$Customer - $BackupType report"
		$MessageAttachement = [Convert]::ToBase64String([IO.File]::ReadAllBytes($TranscriptFile))
		$TranscriptContent = Get-Content $TranscriptFile
		
		#Check for any error file generated by the script
		$ErrorEntries = @()
		foreach($entry in $TranscriptContent) {
			$ErrorEntries += $entry | ?{$_ -match "file://"} 
		}
	
		if(($null -ne $ErrorEntries) -and ($ErrorEntries)) {
			$ErrorLogFile = (($ErrorEntries -split'{|}') -replace ('file://','') -replace ('\/','\') | group | select -ExpandProperty Name)[1]
			if(Test-Path $ErrorLogFile) {
				if($ErrorLogFile -match "\\") {
					Write-Host "[INFO]: Error file '$ErrorLogFile' found to be added to the logs" -ForeGroundColor Yellow
					$MessageAttachement2 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($ErrorLogFile))
						
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
							Attachments = @(
								@{
									"@odata.type" = "#microsoft.graph.fileAttachment"
									Name = ($TranscriptFile -split"\\")[-1]
									ContentType = "text/plain"
									ContentBytes = $MessageAttachement
								}
								@{
									"@odata.type" = "#microsoft.graph.fileAttachment"
									Name = ($ErrorLogFile -split"\\")[-1]
									ContentType = "text/plain"
									ContentBytes = $MessageAttachement2
								}
							)
						}

					SaveToSentItems = "false"
					}
				}
			} else {
				Write-Host "[ERROR]: Error file '$ErrorLogFile' could not be found" -ForeGroundColor Yellow
			}
		} else {
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
					Attachments = @(
						@{
							"@odata.type" = "#microsoft.graph.fileAttachment"
							Name = ($TranscriptFile -split"\\")[-1]
							ContentType = "text/plain"
							ContentBytes = $MessageAttachement
						}
					)
				}

			SaveToSentItems = "false"
			}
		}
		
		Write-Host "[INFO]: Sending email From:'$Sender' To:'$Recipient','$Recipient2' Cc:'$Cc' with Subject:'$Subject'" -ForeGroundColor Yellow
		
		#Connect to Graph by the end of the script to avoid any disconnection during process
		try {
			Connect-MgGraph -ClientId $AppIdExport -Certificate $cert -TenantId $TenantName -ErrorAction Stop | Out-Null
			"[OK] : Connected to Microsoft Graph"
		} catch {
			"[ERROR] : Fail to run cmdlet 'Connect-MgGraph' : {0}" -f $_.Exception[0].Message
			break
		}
		
		try {
			Send-MgUserMail -UserId $Sender -BodyParameter $SendMailParams -ErrorAction Stop
			"[OK] : Email sent From:{0} - To:{1},{2} - Cc:{3} - Subject:{4}" -f $Sender,$Recipient,$Recipient2,$Cc,$Subject
		} catch {
			"[ERROR]: Report could not be sent by mail : {0}" -f $_.Exception[0].Message
		}
	}
}

#Remove logs and backups older than 2 months
try { 
	Get-ChildItem $ScriptLocation -Recurse | ?{$_.Name -notin $DoNotRemoveExceptions -and ($_.LastWriteTime -lt (Get-Date).AddMonths(-2))} | Remove-Item -Recurse -Force -ErrorAction Stop -Verbose
	#Retrieve empty folders to remove them
	$Folders = (Get-ChildItem $ScriptLocation | ?{$_.PsIsContainer}).FullName
	foreach($Folder in $Folders) {
		if((Get-ChildItem $Folder).count -eq 0) {
			#Remove folder if empty
			$Folder | Remove-Item -Recurse -Force -ErrorAction Stop -Verbose
		} else {
			#Folder contains files < 2 months, do not remove
		}
	}
	
} catch {
	Write-Host "[ERROR]: Error logs could not be removed" -ForeGroundColor Yellow
	$_.Exception[0].Message
}

Stop-Transcript