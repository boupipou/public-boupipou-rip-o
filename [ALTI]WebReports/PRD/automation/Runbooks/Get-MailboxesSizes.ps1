param (
	[bool]$DebugEnabled = $false
)

#Troubleshooting/evolution purpose
if($DebugEnabled) {
	$Debug = $true
} else { 
	$Debug = $false 
}
if($Debug) { Write-Output "[DEBUG] : Debug : $Debug" }

<#Add Exchange permissions to Automation account

#>

$JobStart = Get-Date
Write-Output "Jobstart : $($JobStart)"

#Force runbook to use correct modules
Import-Module Az.Accounts
Import-Module Az.Storage
Import-Module Az.Resources
Import-Module ExchangeOnlineManagement

#Debug Step: Log loaded assemblies
if($Debug) { Write-Output "[DEBUG] : $([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -like `"*Azure*`" } | Select FullName, Location)" }

#Export location
$ScriptLocation = $env:TEMP	

#$Script name
$ScriptName = "Get-MailboxesSizes"

$Data = @()
$Report = @()
$Customer = "contoso"

$Date = [datetime]::UtcNow
#France timezone
$franceTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Romance Standard Time")
#Convert date to local time 
$franceNow = [System.TimeZoneInfo]::ConvertTimeFromUtc($Date, $franceTimeZone)
$FormattedDate = $franceNow.ToString("dd_MM_yyyy_HH-mm-ss")

#Retrieve Automation variables
$RGName = Get-AutomationVariable -Name 'RGName'
$StorageAccountName = Get-AutomationVariable -Name 'StorageAccountName'
$StorageAccountContainer = Get-AutomationVariable -Name 'StorageContainerPriv'
$AutomationSubscriptionId = Get-AutomationVariable -Name 'AutomationSubscriptionId'

#Format files names and path for export : temporary Automation storage and move to storage account
#Report file : csv
$ReportFile = "$ScriptName`_$FormattedDate.csv"
$ReportFileCurrent = "$ScriptName`_latest.csv"
$ReportPath = Join-Path $env:TEMP $ReportFile
$BlobNameReport = "data/$ScriptName/$ReportFile"
$BlobNameReportCurrent = "data/$ScriptName/$ReportFileCurrent"

#Connection to Azure Storage using managed identity
try {
    Connect-AzAccount -Identity -ErrorAction Stop
    if($Debug) { Write-Output "[DEBUG] : Connected to Az service" }
} catch {
    Write-Output "[ERROR] : Fail to connect to Az service : $($_.Exception[0].Message)"
	break
}

#Entra application to log on
$ApplicationID = "0a4f8db8-b596-4b79-b30e-dee35456a934" #Azure-MailboxesSizes app
$TenantName = "altiservice.com"
$cert = Get-AutomationCertificate -Name 'M365DSCEncryptionCert' #pfx
#Exchange.ManageAsApp + Global Reader

try {
	Connect-ExchangeOnline -AppId $ApplicationID -Organization $TenantName -Certificate $cert -ErrorAction Stop
	if($Debug) { Write-Output "[OK] : Connected to Exchange Online" }
} catch {
	Write-Output "[ERROR] : Fail to run cmdlet 'Connect-ExchangeOnline' : $($_.Exception[0].Message)"
	break
}

#Script beginning 

#Retrieve all mailboxes
try {
	$Mailboxes = Get-ExoMailbox -ResultSize Unlimited -Properties Guid,UserPrincipalName,ArchiveStatus,RecoverableItemsQuota,ProhibitSendReceiveQuota,UsageLocation,RecipientTypeDetails,AutoExpandingArchiveEnabled -ErrorAction Stop
	if(!([string]::IsNullOrEmpty($Mailboxes))) {
		if($Debug) { Write-Output "[VERBOSE] : Cmdlet Get-Mailbox successfull" }
		
		$TotalMailboxesCount = $Mailboxes | measure | select -ExpandProperty count
		if($Debug) { Write-Output "[VERBOSE] : Total mailboxes : $TotalMailboxesCount" }
	}
} catch {
	Write-Output "[ERROR] : Fail to run cmdlet 'Get-Mailbox' : $($_.Exception[0].Message)"
	break
}

foreach($Mailbox in $Mailboxes) {

	#USER DOMAIN
	if([string]::IsNullOrEmpty($Mailbox.UserPrincipalName)) {
		$UserDomain = ""
	} else {
		$UserDomain = ($Mailbox.UserPrincipalName -split "@")[1]
	}
	if($Debug) { Write-Output "[VERBOSE] : UserDomain : $UserDomain" }
	
	#MAILBOX STATISTICS
	try {
		$MailboxStatistics = Get-EXOMailboxStatistics $Mailbox.Guid.Guid -ErrorAction Stop
		if($Debug) { Write-Output "[VERBOSE] : Cmdlet 'Get-EXOMailboxStatistics' successfull" }
	} catch {
		$MailboxStatistics = ""
		if($Debug) { Write-Output "[ERROR] : Fail to run cmdlet 'Get-EXOMailboxStatistics' for user $($Mailbox.UserPrincipalName) : $($_.Exception[0].Message)" } 
	}
	
	#RECOVERABLE ITEMS STATISTICS
	try {
		$MailboxFolderStatistics = Get-EXOMailboxFolderStatistics $Mailbox.Guid.Guid -FolderScope RecoverableItems -ErrorAction Stop | ?{$_.Name -eq "Recoverable Items"}
		if($Debug) { Write-Output "[VERBOSE] : Cmdlet 'Get-EXOMailboxFolderStatistics' successfull" }
	} catch {
		$MailboxFolderStatistics = ""
		if($Debug) { Write-Output "[ERROR] : Fail to run cmdlet 'Get-EXOMailboxFolderStatistics' for user $($Mailbox.UserPrincipalName) : $($_.Exception[0].Message)" }
	}
	
	#ARCHIVE STATISTICS
	if($Mailbox.ArchiveStatus -eq "Active") {
		try {
			$ArchiveStatistics = Get-EXOMailboxStatistics $Mailbox.Guid.Guid -Archive -ErrorAction Stop
			if($Debug) { Write-Output "[VERBOSE] : Cmdlet 'Get-EXOMailboxStatistics -Archive' successfull" }
		} catch {
			$ArchiveStatistics = "Error"
			if($Debug) { Write-Output "[ERROR] : Fail to run cmdlet 'Get-EXOMailboxStatistics -Archive' for user $($Mailbox.UserPrincipalName) : $($_.Exception[0].Message)" }
		}
	} else {
		if($Debug) { Write-Output "[VERBOSE] : No archive for user $($Mailbox.UserPrincipalName)" }
        $ArchiveStatistics = "None"
	}
	
	if([string]::IsNullOrEmpty($MailboxStatistics)) {
		#No mailbox statistics, cannot retrieve free and occupied space
		$MailboxSize = "Unknown"
		if($Debug) { Write-Output "[WARNING] : MailboxSize : Unknown" }
	} else { 
		#MAILBOX
		#Calculate mailbox size, quota, percent free and threshold
		$MailboxSize = ($MailboxStatistics.TotalItemSize.Value).ToGb()
		if($Debug) { Write-Output "[VERBOSE] : MailboxSize retrieved : $MailboxSize GB" }
		
		$MailboxSizeQuota = [math]::Round(($Mailbox.ProhibitSendReceiveQuota -replace '^.+\((.+\))','$1' -replace '\D' -as [int64])/1GB)
		if($Debug) { Write-Output "[VERBOSE] : MailboxSizeQuota retrieved : $MailboxSizeQuota GB" }
		
		$MailboxSizePercentFree = (($MailboxSizeQuota-$MailboxSize) * 100)/$MailboxSizeQuota
		if($Debug) { Write-Output "[VERBOSE] : MailboxSizePercentFree retrieved : $MailboxSizePercentFree %" }
		
		#Set current mailbox size indicator : critical or warning
		if($MailboxSizePercentFree -le 5) {
			$MailboxSizeThreshold = "Critical"
		} elseif(($MailboxSizePercentFree -le 10) -and ($MailboxSizePercentFree -ge 6)) {
			$MailboxSizeThreshold = "Warning"
		} else {
			$MailboxSizeThreshold = "Ok"
		}		
		if($Debug){ Write-Output "[VERBOSE] : Mailbox free size : $MailboxSizePercentFree % - threshold $MailboxSizeThreshold " }
		
		#DELETED ITEMS
		#Calculate deleted items size, quota, percent free and threshold
		$DeletedItemsSize = ($MailboxStatistics.TotalDeletedItemSize.Value).ToGb()
		if($Debug) { Write-Output "[VERBOSE] : DeletedItemsSize retrieved : $DeletedItemsSize GB" }
		
		$DeletedItemsQuota = [math]::Round(($Mailbox.ProhibitSendReceiveQuota -replace '^.+\((.+\))','$1' -replace '\D' -as [int64])/1GB)
		if($Debug) { Write-Output "[VERBOSE] : DeletedItemsQuota retrieved : $DeletedItemsQuota GB" }
		
		$DeletedItemsPercentFree = (($DeletedItemsQuota-$DeletedItemsSize) * 100)/$DeletedItemsQuota
		if($Debug) { Write-Output "[VERBOSE] : DeletedItemsPercentFree retrieved : $DeletedItemsPercentFree %" }
		
		#Set current deleted items size indicator : critical or warning
		if($DeletedItemsPercentFree -le 5) {
			$DeletedItemsThreshold = "Critical"
		} elseif(($DeletedItemsPercentFree -le 10) -and ($DeletedItemsPercentFree -ge 6)) {
			$DeletedItemsThreshold = "Warning"
		} else {
			$DeletedItemsThreshold = "Ok"
		}
		
		if($Debug) { Write-Output "[VERBOSE] : Deleted items free size : $DeletedItemsPercentFree % - threshold $DeletedItemsThreshold" }
	}
	
	#RECOVERABLE ITEMS
	#Calculate recoverable items size, quota, percent free and threshold
	if([string]::IsNullOrEmpty($MailboxFolderStatistics)) {
		$RecoverableItemsSize = "Unknown"
		if($Debug) { Write-Output "[WARNING] : RecoverableItemsSize : Unknown" }
	} else {
		$RecoverableItemsSize = [math]::Round((($MailboxFolderStatistics | select -ExpandProperty FolderAndSubfolderSize) -replace '^.+\((.+\))','$1' -replace '\D' -as [int64])/1GB)#Reformat 105 MB (110,093,300 bytes) to 110093300 and divide by 1
		if($Debug) { Write-Output "[VERBOSE] : RecoverableItemsSize retrieved : $RecoverableItemsSize GB" }
		
		$RecoverableItemsQuota = $Mailbox.RecoverableItemsQuota.ToString()
        if ($quotaStr -match '\(([0-9,]+)\s*bytes\)') {
            #Remove commas, then convert to GB
            $bytes = ($matches[1] -replace ',','')
            $RecoverableItemsQuota = [math]::Round([double]$bytes / 1GB)
        } elseif ($quotaStr -eq "Unlimited") {
            $RecoverableItemsQuota = "Unlimited"
        } else {
            $RecoverableItemsQuota = 0
        }
		if($Debug) { Write-Output "[VERBOSE] : RecoverableItemsQuota retrieved : $RecoverableItemsQuota GB" }
		
		$RecoverableItemsPercentFree = (($RecoverableItemsQuota-$RecoverableItemsSize) * 100)/$RecoverableItemsQuota
		if($Debug) { Write-Output "[VERBOSE] : RecoverableItemsPercentFree retrieved : $RecoverableItemsPercentFree %" }
		
		#Set current recoverable items size indicator : critical or warning
		if($RecoverableItemsPercentFree -le 5) {
			$RecoverableItemsSizeThreshold = "Critical"
		} elseif(($RecoverableItemsPercentFree -le 10) -and ($RecoverableItemsPercentFree -ge 6)) {
			$RecoverableItemsSizeThreshold = "Warning"
		} else {
			$RecoverableItemsSizeThreshold = "Ok"
		}
		if($Debug) { Write-Output "[VERBOSE] : Recoverable items free size : $RecoverableItemsPercentFree % - threshold $RecoverableItemsSizeThreshold" }
		
		#ARCHIVE
		if([string]::IsNullOrEmpty($ArchiveStatistics)) {
			$ArchiveSize = "Unknown"
			if($Debug) { Write-Output "[WARNING] : ArchiveSize : Unknown" }
		} else {
			$ArchiveSize = ($ArchiveStatistics.TotalItemSize.Value).ToGb()
			if($Debug) { Write-Output "[VERBOSE] : ArchiveSize : $ArchiveSize GB" }
		}
	}
		
	$obj = New-Object System.Object
	$obj | Add-Member -MemberType NoteProperty -Name "UsageLocation" -Value $Mailbox.UsageLocation
	$obj | Add-Member -MemberType NoteProperty -Name "Domain" -Value $UserDomain
	$obj | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $Mailbox.UserPrincipalName
	$obj | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $Mailbox.RecipientTypeDetails
	
	$obj | Add-Member -MemberType NoteProperty -Name "MailboxSize" -Value $MailboxSize
	$obj | Add-Member -MemberType NoteProperty -Name "MailboxSizeQuota" -Value $MailboxSizeQuota
	$obj | Add-Member -MemberType NoteProperty -Name "MailboxSizePercentFree" -Value "$MailboxSizePercentFree%"
	$obj | Add-Member -MemberType NoteProperty -Name "MailboxSizeThreshold" -Value $MailboxSizeThreshold
	
	$obj | Add-Member -MemberType NoteProperty -Name "RecoverableItemsSize" -Value $RecoverableItemsSize
	$obj | Add-Member -MemberType NoteProperty -Name "RecoverableItemsQuota" -Value $RecoverableItemsQuota
	$obj | Add-Member -MemberType NoteProperty -Name "RecoverableItemsPercentFree" -Value "$RecoverableItemsPercentFree%"
	$obj | Add-Member -MemberType NoteProperty -Name "RecoverableItemsSizeThreshold" -Value $RecoverableItemsSizeThreshold
	
	$obj | Add-Member -MemberType NoteProperty -Name "DeletedItemsSize" -Value $DeletedItemsSize
	$obj | Add-Member -MemberType NoteProperty -Name "DeletedItemsQuota" -Value $DeletedItemsQuota
	$obj | Add-Member -MemberType NoteProperty -Name "DeletedItemsPercentFree" -Value "$DeletedItemsPercentFree%"
	$obj | Add-Member -MemberType NoteProperty -Name "DeletedItemsThreshold" -Value $DeletedItemsThreshold
	
	$obj | Add-Member -MemberType NoteProperty -Name "ArchiveSize" -Value $ArchiveSize
	$obj | Add-Member -MemberType NoteProperty -Name "AutoExpandingArchiveEnabled" -Value $Mailbox.AutoExpandingArchiveEnabled	
	
	$Data += $obj
	
	Clear-Variable -Name MailboxSize
	Clear-Variable -Name MailboxSizeQuota
	Clear-Variable -Name MailboxSizePercentFree
	Clear-Variable -Name MailboxSizeThreshold
	Clear-Variable -Name RecoverableItemsSize
	Clear-Variable -Name RecoverableItemsQuota
	Clear-Variable -Name RecoverableItemsPercentFree
	Clear-Variable -Name RecoverableItemsSizeThreshold
	Clear-Variable -Name DeletedItemsSize
	Clear-Variable -Name DeletedItemsQuota
	Clear-Variable -Name DeletedItemsPercentFree
	Clear-Variable -Name DeletedItemsThreshold
	Clear-Variable -Name ArchiveSize
}

#Export data to csv
if(!([string]::IsNullOrEmpty($Data))) {
	try {
		#Export data in a csv file
		$Data | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding Default -Delimiter ";" -ErrorAction Stop
		if($Debug) { Write-Output "[DEBUG] : Export-Csv successfull" }
	} catch {
		Write-Output "[ERROR] : Fail to run cmdlet 'Export-Csv' : $($_.Exception[0].Message)"
	}

    #Operate in the same subscription than automation and storage accounts
    try {
        Set-AzContext -SubscriptionId $AutomationSubscriptionId -ErrorAction Stop
        if($Debug) { Write-Output "[DEBUG] : Set-AzContext successfull" }
    } catch {
        Write-Output "[ERROR] : Fail to run cmdlet 'Set-AzContext' : $($_.Exception[0].Message)"
    }

    #Store report file to storage container
    try {
        $ctx = (Get-AzStorageAccount -ResourceGroupName $RGName -Name $StorageAccountName).Context

        Set-AzStorageBlobContent -File $ReportPath -Container $StorageAccountContainer -Blob $BlobNameReport -Context $ctx -Force -ErrorAction Stop
        Write-Output "[OK] : File $ReportPath successfully created or updated to blob : $BlobNameReport"

        #Create the latest csv file used to load the data in the web report
        Set-AzStorageBlobContent -File $ReportPath -Container $StorageAccountContainer -Blob $BlobNameReportCurrent -Context $ctx -Force -ErrorAction Stop
        Write-Output "[OK] : File $ReportPath successfully created or updated to blob : $BlobNameReportCurrent"
            
    } catch {
        Write-Output "[ERROR] : Fail to run cmdlet 'Set-AzStorageBlobContent' : $($_.Exception[0].Message)"
    }
}

$JobEnd = Get-Date
Write-Output "JobEnd : $($JobEnd)"

$JobDuration = "$([math]::Round($(New-TimeSpan -Start $JobStart -End $JobEnd | select -ExpandProperty TotalMinutes)))"
Write-Output "JobDuration : $($JobDuration)  mn"