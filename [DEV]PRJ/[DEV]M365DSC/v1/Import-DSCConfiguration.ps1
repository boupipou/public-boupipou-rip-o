[CmdletBinding()]
param(
	[Parameter(Mandatory=$false)]
	[string]$PathToDirectoryContainingRestorationFile,
	[Parameter(Mandatory=$false)]
	[switch]$Help=$false
)

$Date = Get-Date
$FormattedDate = $Date.ToString("dd_MM_yyyy_HH-mm-ss")
$TenantId = "gifi.onmicrosoft.com"
#Azure AD application 'Read Only'
$ApplicationIdExport = "733fa52d-7d69-4b40-82af-8b55ba6de454"

#Azure AD application 'Write'
$ApplicationIdImport = "8af3fb62-7ac7-45d1-84a2-a03494e08961"

#Script location retrieved automatically
$ScriptLocation = $PSScriptRoot
#If no script location can be retrieved automatically, script location will be current location
if(!($PSScriptRoot -match "\\")) { $ScriptLocation = Get-Location }

#Script name retrieved automatically
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0]
#If no script name can be retrieved automatically, script name will be defined as below
if(!($MyInvocation.MyCommand.Name -match "\w")) { $ScriptName = "Import-DSCConfiguration" } 

$TranscriptFile = "$ScriptLocation\Transcript_Restoration_$FormattedDate.log"

#Import certificate under local computer certificate store
$cert = Get-ChildItem Cert:\LocalMachine\My\ | ?{$_.Subject.StartsWith("CN=M365DSC")}

#Display help about script execution
switch($Help) {
	$true {
		Clear-Host
		
		Write-Host "[INFO]: This script can run one parameter which is" -ForeGroundColor Gray -NoNewLine;Write-Host " mandatory" -ForeGroundColor Red
		Write-Host "[INFO]: To perform a restoration from a file : .\PathToScript\Import-DSCConfiguration.ps1" -ForeGroundColor Gray -NoNewLine;Write-Host " -PathToDirectoryContainingRestorationFile DiskDrive:\Folder\" -ForeGroundColor Yellow
		Write-Host "Example given: `nC:\Scripts\DSC\" -ForeGroundColor Gray -NoNewLine;Write-Host "FullBackup_20240402\FullBackup\" -ForeGroundColor Yellow
		Write-Host "C:\Scripts\DSC\" -ForeGroundColor Gray -NoNewLine;Write-Host "Exchange_20240402\EXOAcceptedDomain\" -ForeGroundColor Yellow
		Write-Host "C:\Scripts\DSC\" -ForeGroundColor Gray -NoNewLine;Write-Host "customComponents_20240402\EXOTransportRule\" -ForeGroundColor Yellow
		""
		Write-Host "[INFO]: Prior restoring a save file, you will need to execute the .ps1 file in order to generate the compiled 'localhost.mof' file" -ForeGroundColor Gray
		""

		exit
	}
	default {
		#Exit the loop and start the script in normal execution
		break
	}
}

Start-Transcript -Path $TranscriptFile -Force

#Test if a directory has been provided for the restoration
if(!($PathToDirectoryContainingRestorationFile -match "\\")) {
	Write-Host "[INFO]: No directory provided. Please provide a directory path containing a compiled .mof file to restore." -ForeGroundColor Yellow
	""
	
	exit
} else {
	Write-Host "[INFO]: Directory provided : $PathToDirectoryContainingRestorationFile" -ForeGroundColor Yellow
	
	#Test if directory exists
	if(!(Test-Path $PathToDirectoryContainingRestorationFile)) {
		Write-Host "[INFO]: Path to directory provided does not exist. Please provide a directory path containing a compiled .mof file to restore." -ForeGroundColor Yellow
		""
	
		exit
	} else {
		Write-Host "[INFO]: Path to directory exists" -ForeGroundColor Yellow
		#Test if directory contains a compiled file
		try {
			$DirectoryFiles = Get-ChildItem $PathToDirectoryContainingRestorationFile -Filter "localhost.mof" -ErrorAction Stop
		} catch {
			Write-Host "[ERROR]: Files could not be retrieved from '$PathToDirectoryContainingRestorationFile' : $($_.Exception[0].Message)" -ForeGroundColor Yellow
			""
			
			exit
		}
		if($null -ne $DirectoryFiles) {
			Write-Host "[INFO]: Directory contains at least one localhost.mof file" -ForeGroundColor Yellow
			#Directory contains at least one .mof file
			#Has to contain only ONE .mof file
			if($DirectoryFiles.count -gt 1) {
				Write-Host "[ERROR]: More than one compiled .mof file has been found in the directory provided. Move or remove the unwanted files." -ForeGroundColor Yellow
				""
				
				exit
			}
		} else {
			Write-Host "[ERROR]: Directory provided does not contain any localhost.mof file" -ForeGroundColor Yellow
			""
			
			exit
		}
	}
}

#Last write time from the compiled file is retrieved and will be replaced to its original time after the file has been modified
$LastWriteTimeMof = Get-Item "$PathToDirectoryContainingRestorationFile\localhost.mof" | select -ExpandProperty LastWriteTime 
Write-Host "[INFO]: LastWriteTime for localhost.mof file : $LastWriteTimeMof" -ForeGroundColor Yellow

$DSCVersion = Get-InstalledModule Microsoft365DSC | select -ExpandProperty Version

#Replace 'Export AppId' by 'Import AppId' in order to provide the correct permission to 'Write'
#Replace the current compiled file by a temporary file in order to replace it
try {
	(Get-Content "$PathToDirectoryContainingRestorationFile\localhost.mof") -Replace($ApplicationIdExport,$ApplicationIdImport) -Replace(' ModuleVersion = "\d*.\d*.\d*.\d*";'," ModuleVersion = `"$DSCVersion`";") | Set-Content "$PathToDirectoryContainingRestorationFile\localhost2.mof" -ErrorAction Stop
} catch {
	Write-Host "[ERROR]: Could not replace AppId in localhost.mof : $(($_.Exception[0]).Message)" -ForeGroundColor Yellow
	""
	
	exit
}

#Check if the temporary file has been created. If so the original compiled file is removed, and the new compiled .mof file is renamed
if(Test-Path "$PathToDirectoryContainingRestorationFile\localhost2.mof") {
	try {
		Remove-Item "$PathToDirectoryContainingRestorationFile\localhost.mof" -Force -ErrorAction Stop
		Rename-Item "$PathToDirectoryContainingRestorationFile\localhost2.mof" -NewName "localhost.mof" -Force -ErrorAction Stop
	} catch {
		Write-Host "[ERROR]: Could not remove '$PathToDirectoryContainingRestorationFile\localhost.mof' : $($_.Exception[0].Message)" -ForeGroundColor Yellow
		exit
	}
}

#On remet la date de dernière modification originale sur le fichier à restaurer
(Get-Item $PathToDirectoryContainingRestorationFile\localhost.mof).LastWriteTime = $LastWriteTimeMof

#On exécute le process de restauration de la configuration
Start-DSCConfiguration -Path $PathToDirectoryContainingRestorationFile -Wait -Verbose -Force

Stop-Transcript