param (
	[bool]$DebugEnabled = $false
)

#Troubleshooting/evolution purpose
if($DebugEnabled) {
	$Debug = $true
} else { 
	$Debug = $false 
}
if($Debug) { "[DEBUG] : Debug : $Debug" }

#Force runbook to use correct modules
Import-Module Az.Accounts
Import-Module Az.Storage
Import-Module Az.Resources

#Debug Step: Log loaded assemblies
if($Debug) { "[DEBUG] : $([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -like `"*Azure*`" } | Select FullName, Location)" }

#Export location
$ScriptLocation = $env:TEMP	

#$Script name
$ScriptName = "Get-AZRoleAssignments"

$Data = @()
$Report = @()
$Customer = "VUSION"

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
    if($Debug) { "[DEBUG] : Connected to Az service" }
} catch {
    "[ERROR] : Fail to connect to Az service : {0}" -f $_.Exception[0].Message
	break
}

#Script beginning 

#Retrieve Azure subscriptions
try {
	$Subscriptions = Get-AzSubscription -ErrorAction Stop
	if($Debug) { Write-Output "[DEBUG] : AzSubscriptions retrieved" }
} catch {
	Write-Output "[ERROR] : Fail to run cmdlet 'Connect-AzAccount' : $($_.Exception[0].Message)"
	break
}

foreach($Sub in $Subscriptions) {
    Write-Output "--- Subscription : $($Sub.Name) ---"

    #Set context on the current subscription to select it
    try {
		Set-AzContext -SubscriptionId $Sub.Id -ErrorAction Stop
		if($Debug) { Write-Output "[DEBUG] : AzContext retrieved" }
	} catch {
		Write-Output "[ERROR] : Fail to run cmdlet 'Set-AzContext' : $($_.Exception[0].Message)"
		break
	}
	
	#List roles for each subscription
	try {
		$roles = Get-AzRoleAssignment -Scope "/subscriptions/$($Sub.Id)" -ErrorAction Stop #| ?{(!($_.Scope -match "\w+\/providers/+\w"))}
		if($Debug) { Write-Output "[DEBUG] : AzRoleAssignment retrieved for subscription $($Sub.Id)" }
	} catch {
		Write-Output "[ERROR] : Fail to run cmdlet 'Get-AzRoleAssignment' on subscription $($Sub.Id) : $($_.Exception[0].Message)"
	}
	
	foreach ($role in $roles) {
		#Initialize variables for each loop to avoid duplicated values in case of exception
		$ManagementGroup = $null
		$Subscription = $null
		$DisplayName = $null
		$SignInName = $null
		$ObjectType = $null
		$RoleDefinitionName = $null
		$Resource1 = $null
		$Resource2 = $null
		$Scope = $null
		
		#Define data
		if(!([string]::IsNullOrEmpty($Sub.Name))) {
			$Subscription = $Sub.Name
		} else {
			$Subscription = ""
		}
		if($Debug) { Write-Output "[DEBUG] : Subscription : $Subscription" }
		
		if(!([string]::IsNullOrEmpty($role.DisplayName))) {
			$DisplayName = $role.DisplayName
		} else {
			$DisplayName = ""
		}			
		if($Debug) { Write-Output "[DEBUG] : DisplayName : $DisplayName" }

		if(!([string]::IsNullOrEmpty($role.SignInName))) {
			$SignInName = $role.SignInName
		} else {
			$SignInName = ""
		}			
		if($Debug) { Write-Output "[DEBUG] : SignInName : $SignInName" }		

		if(!([string]::IsNullOrEmpty($role.ObjectType))) {
			$ObjectType = $role.ObjectType
		} else {
			$ObjectType = ""
		}			
		if($Debug) { Write-Output "[DEBUG] : ObjectType : $ObjectType" }	

		if(!([string]::IsNullOrEmpty($role.RoleDefinitionName))) {
			$RoleDefinitionName = $role.RoleDefinitionName
		} else {
			$RoleDefinitionName = ""
		}			
		if($Debug) { Write-Output "[DEBUG] : RoleDefinitionName : $RoleDefinitionName" }			


		if(!([string]::IsNullOrEmpty($role.Scope))) {
			$Scope = $role.Scope
		} else {
			$Scope = ""
		}			
		if($Debug) { Write-Output "[DEBUG] : Scope : $Scope" }	

		#If no scope is retrieved, Resource1 and Resource2 are empty
		if($Scope -eq "/") {
			$Resource1 = ""
			$Resource2 = ""
		} else {				
			#Split the last 2 property values from Scope property to format Resource2 : e.g : /storageAccounts/dscdevfrstg
			$Resource2temp = $role | select -ExpandProperty Scope | % { "/$(($_ -split("\/"))[-2])/$(($_ -split("\/"))[-1]) " }
			if(!([string]::IsNullOrEmpty($Resource2temp))) {
				$Resource2 = ($Resource2temp.Replace(" ",""))
				
				#Resource 1 : Scope - Resource2 : e.g : /subscriptions/d7cb1166-a780-4c7b-8435-4605786981eb/resourceGroups/dsc-dev-fr-rg/providers/Microsoft.Storage
				$Resource1 = ($role | select -ExpandProperty Scope) -replace("$Resource2","")
			} else {
				"[ERROR] : Resource2 is null or empty ; cannot format Resource1 and Resource2"
			}
		}
		if($Debug) { Write-Output "[DEBUG] : Resource1 : $Resource1" }
		if($Debug) { Write-Output "[DEBUG] : Resource2 : $Resource2" }
		
		#Retrieve ManagementGroup from scope
		if($Scope.startswith("/providers/Microsoft.Management/managementGroups/", "CurrentCultureIgnoreCase")) {
			#If scope includes ManagementGroup, populate ManagementGroup but not Subscription
			$ManagementGroup = ($Scope -split("/"))[-1] #/providers/Microsoft.Management/managementGroups/supiti_dev => supiti_dev
			$Subscription = ""
		} else {
			#If not, populate only Subscription
			$ManagementGroup = ""
		}
		if($Debug) { Write-Output "[DEBUG] : ManagementGroup : $ManagementGroup" }
		
		$obj = New-Object System.Object
		$obj | Add-Member -MemberType NoteProperty -Name "ManagementGroup" -Value $ManagementGroup
		$obj | Add-Member -MemberType NoteProperty -Name "Subscription" -Value $Subscription
		$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $DisplayName
		$obj | Add-Member -MemberType NoteProperty -Name "SignInName" -Value $SignInName
		$obj | Add-Member -MemberType NoteProperty -Name "ObjectType" -Value $ObjectType
		$obj | Add-Member -MemberType NoteProperty -Name "RoleDefinitionName" -Value $RoleDefinitionName
		$obj | Add-Member -MemberType NoteProperty -Name "Resource1" -Value $Resource1
		$obj | Add-Member -MemberType NoteProperty -Name "Resource2" -Value $Resource2
		$obj | Add-Member -MemberType NoteProperty -Name "Scope" -Value $Scope

		$Data += $obj
	}
}

if(!([string]::IsNullOrEmpty($Data))) {
	try {
		#Export data in a csv file
		$Data | sort ManagementGroup,Subscription,RoleDefinitionName,Resource2 | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding Default -Delimiter ";" -ErrorAction Stop
		if($Debug) { Write-Output "[DEBUG] : Export csv successfull" }
	} catch {
		"[ERROR] : Fail to run cmdlet 'Export-Csv' : {0}" -f $_.Exception[0].Message
	}
    
    #Operate in the same subscription than automation and storage accounts
    Set-AzContext -SubscriptionId $AutomationSubscriptionId
     
	#Store report file to storage container
    try {
        $ctx = (Get-AzStorageAccount -ResourceGroupName $RGName -Name $StorageAccountName).Context

        Set-AzStorageBlobContent -File $ReportPath -Container $StorageAccountContainer -Blob $BlobNameReport -Context $ctx -Force -ErrorAction Stop
        "[OK] : File {0} successfully created or updated to blob : {1}" -f $ReportPath,$BlobNameReport

		#Create the latest csv file used to load the data in the web report
        Set-AzStorageBlobContent -File $ReportPath -Container $StorageAccountContainer -Blob $BlobNameReportCurrent -Context $ctx -Force -ErrorAction Stop
        "[OK] : File {0} successfully created or updated to blob : {1}" -f $ReportPath,$BlobNameReportCurrent
            
    } catch {
        "[ERROR] : Fail to run cmdlet 'Set-AzStorageBlobContent' : {0}" -f $_.Exception[0].Message
    }
}