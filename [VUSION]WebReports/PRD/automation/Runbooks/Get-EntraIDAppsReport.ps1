param (
	[bool]$GlobalReport = $false,
	[bool]$DebugEnabled = $false
)

#Do not send mail if debug mode is used
#Troubleshooting/evolution purpose
if($DebugEnabled) {
	$Debug = $true
} else { 
	$Debug = $false 
}
if($Debug -eq $true) { "[DEBUG] : Debug : $Debug" }

#If parameter -GlobalReport is specified, make a global app report 
#if not, make a report with expiring soon certificates and secret keys only
if($GlobalReport) {
	$GlobalReport = $true
} else {
	$GlobalReport = $false
}
if($Debug -eq $true) { "[DEBUG] : GlobalReport : $GlobalReport" }

#Force runbook to use correct modules
Import-Module Az.Accounts
Import-Module Az.Storage
Import-Module Az.Resources
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Applications
Import-Module Microsoft.Graph.Identity.SignIns
Import-Module Microsoft.Graph.Users.Actions

#Debug Step: Log loaded assemblies
if($Debug -eq $true) { "[DEBUG] : $([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -like `"*Azure*`" } | Select FullName, Location)" }

#Export location
$ScriptLocation = $env:TEMP					

#$Script name
$ScriptName = "Get-EntraIDAppsReport"

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

#Expiration date warning is set to 30 days
[int]$DaysWithinExpiration = 30

#Azure AD application to log on
$ApplicationID = "10368e52-d3fc-4620-a907-43d6eb75d0a2" #MS365-EntraAppsExpirationMonitoring
$TenantID = "026f3f97-f463-4c41-ba74-bc156c3be494"
$cert = Get-AutomationCertificate -Name 'M365DSCEncryptionCert' #pfx
$Sender = "noreply@vusion.com"
if($Debug -eq $true) { "[DEBUG] : Sender : {0}" -f $Sender }

if($Debug -eq $false) {
	$Recipient = "393e4e28.storeelectronicsystems.onmicrosoft.com@fr.teams.ms"
} else {
	$Recipient = "benjamin.poulain@talan.com"
	"[DEBUG] : Debug recipient : {0}" -f $Recipient
}


#Connection to Microsoft Graph
try {
	Connect-MgGraph -AppId $ApplicationID -TenantId $TenantID -Certificate $cert  -ErrorAction Stop | Out-Null
	if($Debug -eq $true) { "[DEBUG] : Connected to Microsoft Graph" }
} catch {
	"[ERROR] : Fail to connect to Microsoft.Graph : {0}" -f $_.Exception[0].Message
	break
}

#Connection to Azure Storage using managed identity
try {
    Connect-AzAccount -Identity -ErrorAction Stop
    if($Debug -eq $true) { "[DEBUG] : Connected to Az service" }
} catch {
    "[ERROR] : Fail to connect to Az service : {0}" -f $_.Exception[0].Message
	break
}

#Script beginning 
 
try {
    $ApplicationList = Get-MgApplication -All -PageSize 999 -ErrorAction Stop
	
	if(!([string]::IsNullOrEmpty($ApplicationList))) {
		$AppCount = $ApplicationList | measure | select -ExpandProperty count
		if($Debug -eq $true) { "[DEBUG] : Entra Applications retrieved : {0} applications" -f $AppCount }
	} else {
		"[ERROR] : Entra applications null or empty: {0}"
		break
	}
} catch {
	"[ERROR] : Fail to retrieve entra applications : {0}" -f $_.Exception[0].Message
	break
}	

#Apps using certificate
$CertificateApps  = $ApplicationList | Where-Object {$_.KeyCredentials}

foreach ($App in $CertificateApps) {
	"--- Application : {0} - {1} ---" -f $App.DisplayName,$App.AppId
    foreach ($Cert in $App.KeyCredentials) {
		$Owners = @()
		$ApplicationPermission = $null
		$ApplicationPermissions = @()
		
		$DaysUntilExpiration = [math]::round((($Cert.EndDateTime) - (Get-Date)).TotalDays)
		
		#Retrieve application service principal name
		try {
			$SPN = (Get-MgServicePrincipal -Filter "AppId eq '$($App.AppId)'" -ErrorAction Stop).Id
			"SPN : $SPN"
		} catch { 
			$SPN = ""
			"[ERROR] : No SPN could be retrieved for this application : {0}" -f $_.Exception[0].Message
		}
		
		#Retrieve application users and groups
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$UsersAndGroups = (Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $SPN -ErrorAction Stop).PrincipalDisplayName -join ";"
				if(!([string]::IsNullOrEmpty($UsersAndGroups))) {
					"UsersAndGroups : $UsersAndGroups"
				} else { 
					"UsersAndGroups : No user or group"
				}
			} catch {
				$UsersAndGroups = ""
				"[ERROR] : No users and groups could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"[ERROR] : SPN null or empty : : no users and groups could be retrieved for this application"
		}


		#Retrieve application "Delegated" permissions
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$DelegatedPermissions = (Get-MgOauth2PermissionGrant -All -ErrorAction Stop | Where-Object {$_.ClientId -eq $SPN}).Scope -split " " -join ";"
				if(!([string]::IsNullOrEmpty($DelegatedPermissions))) {
					"DelegatedPermissions : $DelegatedPermissions"
				} else { 
					"DelegatedPermissions : No delegated permission"
				}
			} catch {
				$DelegatedPermissions = ""
				"[ERROR] : No delegated permissions could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"[ERROR] : SPN null or empty : no delegated permissions could be retrieved for this application"
		}


		#Retrieve application "Application" permissions
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$ApplicationPermissionsAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SPN -ErrorAction Stop
				if(!([string]::IsNullOrEmpty($ApplicationPermissionsAssignment))) {
					Foreach($Appli in $ApplicationPermissionsAssignment) {
						$ApplicationPermissionsResource = $Appli.ResourceDisplayName
						$ApplicationPermissionsId = $Appli.AppRoleId
						$ApplicationPermissionValue = (Get-MgServicePrincipal -Filter "displayName eq '$ApplicationPermissionsResource'" | select -ExpandProperty AppRoles | ?{$_.Id -eq $ApplicationPermissionsId}).Value
						$ApplicationPermission = "{0}:{1}" -f $ApplicationPermissionsResource,$ApplicationPermissionValue
						$ApplicationPermissions += [array]$ApplicationPermission
					}
					"ApplicationPermissions : $ApplicationPermissions"
					
				} else {
					"ApplicationPermissions : No application permission"
				}
			} catch {
				$ApplicationPermissions = ""
				"[ERROR] : No application permissions could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"[ERROR] : SPN null or empty : no application permissions could be retrieved for this application"
		}

		
		#App owners
		try {
			$Owners = (Get-MgApplicationOwner -ApplicationId $App.Id -ErrorAction Stop).AdditionalProperties.userPrincipalName -join ";"
			"Owners : $Owners"
		} catch {
			$Owners = ""
			"[ERROR] : No owners could be retrieved for this application : {0}" -f $_.Exception[0].Message
		}
		
		#Certificate thumbprint
		$ThumbPrint = [System.Convert]::ToBase64String($Cert.CustomKeyIdentifier)
		
		#Certificate expired
		if ($Cert.EndDateTime -lt (Get-Date)) {
			$Expired = $true
			"Certificate expired"
		} else {
			#Certificate not expired
			$Expired = $false
			"Certificate not expired"
		} 
		
		#Certificate expire soon (current expiration date <= today + $DaysWithinExpiration)
        if ($Cert.EndDateTime -le (Get-Date).AddDays($DaysWithinExpiration)) {
			$ExpireSoon = $true
			"Certificate expire soon"
		} else {
			#Certificate does not expire soon
			$ExpireSoon = $false
			"Certificate does not expire soon"
		}
		
		#Sign in logs : interactive, non interactive, service principal, managed identities
		<#$LastInteractiveAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'interactiveUser')" -Top 1
		$LastNonInteractiveAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'nonInteractiveUser')" -Top 1
		$LastServicePrincipalAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'servicePrincipal')" -Top 1
		$LastManagedIdentityAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'managedIdentity')" -Top 1 #>
		
		#Certificate expired or expire soon
		if($Expired -eq $true -or $ExpireSoon -eq $true) {
			if($GlobalReport -eq $false) {
				$AppData = New-Object System.Object
				$AppData | Add-Member -MemberType NoteProperty -Name "AppDisplayName" -Value $App.DisplayName
				$AppData | Add-Member -MemberType NoteProperty -Name "AppId" -Value $App.AppId	
				$AppData | Add-Member -MemberType NoteProperty -Name "KeyType" -Value "Certificate"
				$AppData | Add-Member -MemberType NoteProperty -Name "ExpirationDate" -Value $Cert.EndDateTime 
				$AppData | Add-Member -MemberType NoteProperty -Name "DaysUntilExpiration" -Value $DaysUntilExpiration
				$AppData | Add-Member -MemberType NoteProperty -Name "Expired" -Value $Expired
				$AppData | Add-Member -MemberType NoteProperty -Name "ExpireSoon" -Value $ExpireSoon
				$AppData | Add-Member -MemberType NoteProperty -Name "Owners" -Value $Owners
				$AppData | Add-Member -MemberType NoteProperty -Name "UsersAndGroups" -Value $UsersAndGroups
				$AppData | Add-Member -MemberType NoteProperty -Name "DelegatedPermissions" -Value $DelegatedPermissions
				$AppData | Add-Member -MemberType NoteProperty -Name "ApplicationPermissions" -Value "$($ApplicationPermissions -join ",")"
				$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalnameId" -Value $SPN
				$AppData | Add-Member -MemberType NoteProperty -Name "ThumbPrint" -Value $ThumbPrint
				$AppData | Add-Member -MemberType NoteProperty -Name "CreatedDateTime" -Value $App.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "Notes" -Value $App.Notes
				
				<#if($null -ne $LastInteractiveAccess) {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value "InteractiveUser"
				} else {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value " "
				}
				$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeInteractiveUser" -Value $LastInteractiveAccess.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameInteractiveUser" -Value $LastInteractiveAccess.UserPrincipalName
				$AppData | Add-Member -MemberType NoteProperty -Name "ClientAppUsedInteractiveUser" -Value $LastInteractiveAccess.ClientAppUsed
				
				if($null -ne $LastNonInteractiveAccess) {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value "NonInteractiveUser"
				} else {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value " "
				}
				$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeNonInteractiveUser" -Value $LastNonInteractiveAccess.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameNonInteractiveUser" -Value $LastNonInteractiveAccess.UserPrincipalName
				$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameNonInteractiveUser" -Value $LastNonInteractiveAccess.ResourceDisplayName
				
				if($null -ne $LastServicePrincipalAccess) {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value "ServicePrincipal"
				} else {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value " "
				}
				$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeServicePrincipal" -Value $LastServicePrincipalAccess.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalNameServicePrincipal" -Value $LastServicePrincipalAccess.ServicePrincipalName
				$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameServicePrincipal" -Value  $LastServicePrincipalAccess.ResourceDisplayName
				
				if($null -ne $LastManagedIdentityAccess) {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value "ManagedIdentity"
				} else {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value " "
				}
				$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeManagedIdentity" -Value $LastManagedIdentityAccess.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityAssociatedResourceId" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.AssociatedResourceId
				$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityMsiType" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.MsiType
				$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameManagedIdentity" -Value $LastManagedIdentityAccess.ResourceDisplayName	#>


				$Data += $AppData
			}
		}
		
		#If global report is requested : retrieve app information
		if($GlobalReport -eq $true) {
			$AppData = New-Object System.Object
			$AppData | Add-Member -MemberType NoteProperty -Name "AppDisplayName" -Value $App.DisplayName
			$AppData | Add-Member -MemberType NoteProperty -Name "AppId" -Value $App.AppId	
			$AppData | Add-Member -MemberType NoteProperty -Name "KeyType" -Value "Certificate"
			$AppData | Add-Member -MemberType NoteProperty -Name "ExpirationDate" -Value $Cert.EndDateTime 
			$AppData | Add-Member -MemberType NoteProperty -Name "DaysUntilExpiration" -Value $DaysUntilExpiration
			$AppData | Add-Member -MemberType NoteProperty -Name "Expired" -Value $Expired
			$AppData | Add-Member -MemberType NoteProperty -Name "ExpireSoon" -Value $ExpireSoon
			$AppData | Add-Member -MemberType NoteProperty -Name "Owners" -Value $Owners
			$AppData | Add-Member -MemberType NoteProperty -Name "UsersAndGroups" -Value $UsersAndGroups
			$AppData | Add-Member -MemberType NoteProperty -Name "DelegatedPermissions" -Value $DelegatedPermissions
			$AppData | Add-Member -MemberType NoteProperty -Name "ApplicationPermissions" -Value "$($ApplicationPermissions -join ",")"
			$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalnameId" -Value $SPN
			$AppData | Add-Member -MemberType NoteProperty -Name "ThumbPrint" -Value $ThumbPrint
			$AppData | Add-Member -MemberType NoteProperty -Name "CreatedDateTime" -Value $App.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "Notes" -Value $App.Notes
			
			<#if($null -ne $LastInteractiveAccess) {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value "InteractiveUser"
			} else {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value " "
			}
			$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeInteractiveUser" -Value $LastInteractiveAccess.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameInteractiveUser" -Value $LastInteractiveAccess.UserPrincipalName
			$AppData | Add-Member -MemberType NoteProperty -Name "ClientAppUsedInteractiveUser" -Value $LastInteractiveAccess.ClientAppUsed
			
			if($null -ne $LastNonInteractiveAccess) {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value "NonInteractiveUser"
			} else {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value " "
			}
			$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeNonInteractiveUser" -Value $LastNonInteractiveAccess.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameNonInteractiveUser" -Value $LastNonInteractiveAccess.UserPrincipalName
			$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameNonInteractiveUser" -Value $LastNonInteractiveAccess.ResourceDisplayName
			
			if($null -ne $LastServicePrincipalAccess) {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value "ServicePrincipal"
			} else {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value " "
			}
			$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeServicePrincipal" -Value $LastServicePrincipalAccess.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalNameServicePrincipal" -Value $LastServicePrincipalAccess.ServicePrincipalName
			$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameServicePrincipal" -Value  $LastServicePrincipalAccess.ResourceDisplayName
			
			if($null -ne $LastManagedIdentityAccess) {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value "ManagedIdentity"
			} else {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value " "
			}
			$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeManagedIdentity" -Value $LastManagedIdentityAccess.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityAssociatedResourceId" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.AssociatedResourceId
			$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityMsiType" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.MsiType
			$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameManagedIdentity" -Value $LastManagedIdentityAccess.ResourceDisplayName	#>

			$Data += $AppData
		}
	}
}

                   

#Apps using client secret
$ClientSecretApps = $ApplicationList | Where-Object {$_.passwordCredentials}

foreach ($App in $ClientSecretApps) {
	"--- Application : {0} - {1} ---" -f $App.DisplayName,$App.AppId
	
    foreach ($Secret in $App.PasswordCredentials) { 
		$Owners = @()
		$ApplicationPermission = $null
		$ApplicationPermissions = @()		
	
		$DaysUntilExpiration = [math]::round((($Secret.EndDateTime) - (Get-Date)).TotalDays)
	
		#Retrieve application service principal name
		try {
			$SPN = (Get-MgServicePrincipal -Filter "AppId eq '$($App.AppId)'" -ErrorAction Stop).Id
			"SPN : $SPN"
		} catch { 
			$SPN = ""
			"[ERROR] : No SPN could be retrieved for this application : {0}" -f $_.Exception[0].Message
		}
		
		#Retrieve application users and groups
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$UsersAndGroups = (Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $SPN -ErrorAction Stop).PrincipalDisplayName -join ";"
				if(!([string]::IsNullOrEmpty($UsersAndGroups))) {
					"UsersAndGroups : $UsersAndGroups"
				} else { 
					"UsersAndGroups : No user or group"
				}
			} catch {
				$UsersAndGroups = ""
				"[ERROR] : No users and groups could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"SPN null or empty : : no users and groups could be retrieved for this application"
		}


		#Retrieve application "Delegated" permissions
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$DelegatedPermissions = (Get-MgOauth2PermissionGrant -All -ErrorAction Stop | Where-Object {$_.ClientId -eq $SPN}).Scope -split " " -join ";"
				if(!([string]::IsNullOrEmpty($DelegatedPermissions))) {
					"DelegatedPermissions : $DelegatedPermissions"
				} else { 
					"DelegatedPermissions : No delegated permission"
				}				
			} catch {
				$DelegatedPermissions = ""
				"[ERROR] : No delegated permissions could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"[ERROR] : SPN null or empty : no delegated permissions could be retrieved for this application"
		}


		#Retrieve application "Application" permissions
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$ApplicationPermissionsAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SPN -ErrorAction Stop
				if(!([string]::IsNullOrEmpty($ApplicationPermissionsAssignment))) {
					Foreach($Appli in $ApplicationPermissionsAssignment) {
						$ApplicationPermissionsResource = $Appli.ResourceDisplayName
						$ApplicationPermissionsId = $Appli.AppRoleId
						$ApplicationPermissionValue = (Get-MgServicePrincipal -Filter "displayName eq '$ApplicationPermissionsResource'" | select -ExpandProperty AppRoles | ?{$_.Id -eq $ApplicationPermissionsId}).Value
						$ApplicationPermission = "{0}:{1}" -f $ApplicationPermissionsResource,$ApplicationPermissionValue
						$ApplicationPermissions += [array]$ApplicationPermission
					}
					"ApplicationPermissions : $ApplicationPermissions"
					
				} else {
					"ApplicationPermissions : No application permission"
				}
			} catch {
				$ApplicationPermissions = ""
				"[ERROR] : No application permissions could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"SPN null or empty : no application permissions could be retrieved for this application"
		}

		
		#App owners
		try {
			$Owners = (Get-MgApplicationOwner -ApplicationId $App.Id -ErrorAction Stop).AdditionalProperties.userPrincipalName -join ";"
			"Owners : $Owners"
		} catch {
			$Owners = ""
			"[ERROR] : No owners could be retrieved for this application : {0}" -f $_.Exception[0].Message
		}

		#Secret expired
		if ($Secret.EndDateTime -lt (Get-Date)) {
			$Expired = $true
			"Secret expired"
		} else {
			#Secret not expired
			$Expired = $false
			"Secret not expired"
		} 
		
		#Secret expire soon (current expiration date <= today + $DaysWithinExpiration)
        if ($Secret.EndDateTime -le (Get-Date).AddDays($DaysWithinExpiration)) {
			$ExpireSoon = $true
			"Secret expire soon"
		} else {
			#Secret does not expire soon
			$ExpireSoon = $false
			"Secret does not expire soon"
		} 
		
		#Apps sign in logs
		<#$LastInteractiveAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'interactiveUser')" -Top 1
		$LastNonInteractiveAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'nonInteractiveUser')" -Top 1
		$LastServicePrincipalAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'servicePrincipal')" -Top 1
		$LastManagedIdentityAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'managedIdentity')" -Top 1 #>

		#Secret expired or expire soon
		if($Expired -eq $true -or $ExpireSoon -eq $true) {
			if($GlobalReport -eq $false) {
				$AppData = New-Object System.Object
				$AppData | Add-Member -MemberType NoteProperty -Name "AppDisplayName" -Value $App.DisplayName
				$AppData | Add-Member -MemberType NoteProperty -Name "AppId" -Value $App.AppId	
				$AppData | Add-Member -MemberType NoteProperty -Name "KeyType" -Value "ClientSecret"
				$AppData | Add-Member -MemberType NoteProperty -Name "ExpirationDate" -Value $Secret.EndDateTime 
				$AppData | Add-Member -MemberType NoteProperty -Name "DaysUntilExpiration" -Value $DaysUntilExpiration 
				$AppData | Add-Member -MemberType NoteProperty -Name "Expired" -Value $Expired
				$AppData | Add-Member -MemberType NoteProperty -Name "ExpireSoon" -Value $ExpireSoon
				$AppData | Add-Member -MemberType NoteProperty -Name "Owners" -Value $Owners
				$AppData | Add-Member -MemberType NoteProperty -Name "UsersAndGroups" -Value $UsersAndGroups
				$AppData | Add-Member -MemberType NoteProperty -Name "DelegatedPermissions" -Value $DelegatedPermissions
				$AppData | Add-Member -MemberType NoteProperty -Name "ApplicationPermissions" -Value "$($ApplicationPermissions -join ",")"
				$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalnameId" -Value $SPN
				$AppData | Add-Member -MemberType NoteProperty -Name "ThumbPrint" -Value ""
				$AppData | Add-Member -MemberType NoteProperty -Name "CreatedDateTime" -Value $App.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "Notes" -Value $App.Notes

				<#if($null -ne $LastInteractiveAccess) {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value "InteractiveUser"
				} else {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value " "
				}
				$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeInteractiveUser" -Value $LastInteractiveAccess.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameInteractiveUser" -Value $LastInteractiveAccess.UserPrincipalName
				$AppData | Add-Member -MemberType NoteProperty -Name "ClientAppUsedInteractiveUser" -Value $LastInteractiveAccess.ClientAppUsed
				
				if($null -ne $LastNonInteractiveAccess) {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value "NonInteractiveUser"
				} else {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value " "
				}
				$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeNonInteractiveUser" -Value $LastNonInteractiveAccess.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameNonInteractiveUser" -Value $LastNonInteractiveAccess.UserPrincipalName
				$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameNonInteractiveUser" -Value $LastNonInteractiveAccess.ResourceDisplayName
				
				if($null -ne $LastServicePrincipalAccess) {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value "ServicePrincipal"
				} else {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value " "
				}
				$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeServicePrincipal" -Value $LastServicePrincipalAccess.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalNameServicePrincipal" -Value $LastServicePrincipalAccess.ServicePrincipalName
				$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameServicePrincipal" -Value  $LastServicePrincipalAccess.ResourceDisplayName
				
				if($null -ne $LastManagedIdentityAccess) {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value "ManagedIdentity"
				} else {
					$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value " "
				}
				$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeManagedIdentity" -Value $LastManagedIdentityAccess.CreatedDateTime
				$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityAssociatedResourceId" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.AssociatedResourceId
				$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityMsiType" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.MsiType
				$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameManagedIdentity" -Value $LastManagedIdentityAccess.ResourceDisplayName #>

				$Data += $AppData			
			}
		}
		
		#If global report is requested : retrieve app information
		if($GlobalReport -eq $true) {
			$AppData = New-Object System.Object
			$AppData | Add-Member -MemberType NoteProperty -Name "AppDisplayName" -Value $App.DisplayName
			$AppData | Add-Member -MemberType NoteProperty -Name "AppId" -Value $App.AppId	
			$AppData | Add-Member -MemberType NoteProperty -Name "KeyType" -Value "ClientSecret"
			$AppData | Add-Member -MemberType NoteProperty -Name "ExpirationDate" -Value $Secret.EndDateTime 
			$AppData | Add-Member -MemberType NoteProperty -Name "DaysUntilExpiration" -Value $DaysUntilExpiration
			$AppData | Add-Member -MemberType NoteProperty -Name "Expired" -Value $Expired
			$AppData | Add-Member -MemberType NoteProperty -Name "ExpireSoon" -Value $ExpireSoon
			$AppData | Add-Member -MemberType NoteProperty -Name "Owners" -Value $Owners
			$AppData | Add-Member -MemberType NoteProperty -Name "UsersAndGroups" -Value $UsersAndGroups
			$AppData | Add-Member -MemberType NoteProperty -Name "DelegatedPermissions" -Value $DelegatedPermissions
			$AppData | Add-Member -MemberType NoteProperty -Name "ApplicationPermissions" -Value "$($ApplicationPermissions -join ",")"
			$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalnameId" -Value $SPN
			$AppData | Add-Member -MemberType NoteProperty -Name "ThumbPrint" -Value ""
			$AppData | Add-Member -MemberType NoteProperty -Name "CreatedDateTime" -Value $App.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "Notes" -Value $App.Notes

			<#if($null -ne $LastInteractiveAccess) {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value "InteractiveUser"
			} else {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value " "
			}
			$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeInteractiveUser" -Value $LastInteractiveAccess.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameInteractiveUser" -Value $LastInteractiveAccess.UserPrincipalName
			$AppData | Add-Member -MemberType NoteProperty -Name "ClientAppUsedInteractiveUser" -Value $LastInteractiveAccess.ClientAppUsed
			
			if($null -ne $LastNonInteractiveAccess) {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value "NonInteractiveUser"
			} else {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value " "
			}
			$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeNonInteractiveUser" -Value $LastNonInteractiveAccess.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameNonInteractiveUser" -Value $LastNonInteractiveAccess.UserPrincipalName
			$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameNonInteractiveUser" -Value $LastNonInteractiveAccess.ResourceDisplayName
			
			if($null -ne $LastServicePrincipalAccess) {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value "ServicePrincipal"
			} else {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value " "
			}
			$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeServicePrincipal" -Value $LastServicePrincipalAccess.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalNameServicePrincipal" -Value $LastServicePrincipalAccess.ServicePrincipalName
			$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameServicePrincipal" -Value  $LastServicePrincipalAccess.ResourceDisplayName
			
			if($null -ne $LastManagedIdentityAccess) {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value "ManagedIdentity"
			} else {
				$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value " "
			}
			$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeManagedIdentity" -Value $LastManagedIdentityAccess.CreatedDateTime
			$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityAssociatedResourceId" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.AssociatedResourceId
			$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityMsiType" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.MsiType
			$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameManagedIdentity" -Value $LastManagedIdentityAccess.ResourceDisplayName	#>	
			
			$Data += $AppData
		}
	}
}



#Apps with no certificate or client secret
if($GlobalReport -eq $true) {
	$AppsWithoutExpiration = $ApplicationList | Where-Object {!($_.passwordCredentials) -and !($_.KeyCredentials)}
	foreach ($App in $AppsWithoutExpiration) {
		"--- Application : {0} - {1} ---" -f $App.DisplayName,$App.AppId
		
		$Owners = @()
		$ApplicationPermission = $null
		$ApplicationPermissions = @()	
		
		#Retrieve application service principal name
		try {
			$SPN = (Get-MgServicePrincipal -Filter "AppId eq '$($App.AppId)'" -ErrorAction Stop).Id
			"SPN : $SPN"
		} catch { 
			$SPN = ""
			"[ERROR] : No SPN could be retrieved for this application : {0}" -f $_.Exception[0].Message
		}
		
		#Retrieve application users and groups
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$UsersAndGroups = (Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $SPN -ErrorAction Stop).PrincipalDisplayName -join ";"
				if(!([string]::IsNullOrEmpty($UsersAndGroups))) {
					"UsersAndGroups : $UsersAndGroups"
				} else { 
					"UsersAndGroups : No user or group"
				}
			} catch {
				$UsersAndGroups = ""
				"[ERROR] : No users and groups could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"SPN null or empty : : no users and groups could be retrieved for this application"
		}


		#Retrieve application "Delegated" permissions
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$DelegatedPermissions = (Get-MgOauth2PermissionGrant -All -ErrorAction Stop | Where-Object {$_.ClientId -eq $SPN}).Scope -split " " -join ";"
				if(!([string]::IsNullOrEmpty($DelegatedPermissions))) {
					"DelegatedPermissions : $DelegatedPermissions"
				} else { 
					"DelegatedPermissions : No delegated permission"
				}				
			} catch {
				$DelegatedPermissions = ""
			"[ERROR] : No delegated permissions could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"[ERROR] : SPN null or empty : no delegated permissions could be retrieved for this application"
		}


		#Retrieve application "Application" permissions
		if(!([string]::IsNullOrEmpty($SPN))) {
			try {
				$ApplicationPermissionsAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $SPN -ErrorAction Stop
				if(!([string]::IsNullOrEmpty($ApplicationPermissionsAssignment))) {
					Foreach($Appli in $ApplicationPermissionsAssignment) {
						$ApplicationPermissionsResource = $Appli.ResourceDisplayName
						$ApplicationPermissionsId = $Appli.AppRoleId
						$ApplicationPermissionValue = (Get-MgServicePrincipal -Filter "displayName eq '$ApplicationPermissionsResource'" | select -ExpandProperty AppRoles | ?{$_.Id -eq $ApplicationPermissionsId}).Value
						$ApplicationPermission = "{0}:{1}" -f $ApplicationPermissionsResource,$ApplicationPermissionValue
						$ApplicationPermissions += [array]$ApplicationPermission
					}
					"ApplicationPermissions : $ApplicationPermissions"
				} else {
					"ApplicationPermissions : No application permission"
				}
			} catch {
				$ApplicationPermissions = ""
				"[ERROR] : No application permissions could be retrieved for this application : {0}" -f $_.Exception[0].Message
			}
		} else {
			"SPN null or empty : no application permissions could be retrieved for this application"
		}

		
		#App owners
		try {
			$Owners = (Get-MgApplicationOwner -ApplicationId $App.Id -ErrorAction Stop).AdditionalProperties.userPrincipalName -join ";"
			"Owners : $Owners"
		} catch {
			$Owners = ""
			"[ERROR] : No owners could be retrieved for this application : {0}" -f $_.Exception[0].Message
		}
  
		<#Apps sign in logs
		$LastInteractiveAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'interactiveUser')" -Top 1
		$LastNonInteractiveAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'nonInteractiveUser')" -Top 1
		$LastServicePrincipalAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'servicePrincipal')" -Top 1
		$LastManagedIdentityAccess = Get-MgBetaAuditLogSignIn -Filter "appid eq '$($App.AppId)' and signInEventTypes/any(t:t eq 'managedIdentity')" -Top 1 #>
			
		$AppData = New-Object System.Object
		$AppData | Add-Member -MemberType NoteProperty -Name "AppDisplayName" -Value $App.DisplayName
		$AppData | Add-Member -MemberType NoteProperty -Name "AppId" -Value $App.AppId	
		$AppData | Add-Member -MemberType NoteProperty -Name "KeyType" -Value ""
		$AppData | Add-Member -MemberType NoteProperty -Name "ExpirationDate" -Value ""
		$AppData | Add-Member -MemberType NoteProperty -Name "DaysUntilExpiration" -Value ""
		$AppData | Add-Member -MemberType NoteProperty -Name "Expired" -Value ""
		$AppData | Add-Member -MemberType NoteProperty -Name "ExpireSoon" -Value ""
		$AppData | Add-Member -MemberType NoteProperty -Name "Owners" -Value $Owners
		$AppData | Add-Member -MemberType NoteProperty -Name "UsersAndGroups" -Value $UsersAndGroups
		$AppData | Add-Member -MemberType NoteProperty -Name "DelegatedPermissions" -Value $DelegatedPermissions
		$AppData | Add-Member -MemberType NoteProperty -Name "ApplicationPermissions" -Value "$($ApplicationPermissions -join ",")"
		$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalnameId" -Value $SPN
		$AppData | Add-Member -MemberType NoteProperty -Name "ThumbPrint" -Value ""
		$AppData | Add-Member -MemberType NoteProperty -Name "CreatedDateTime" -Value $App.CreatedDateTime
		$AppData | Add-Member -MemberType NoteProperty -Name "Notes" -Value $App.Notes

		<# if($null -ne $LastInteractiveAccess) {
			$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value "InteractiveUser"
		} else {
			$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeInteractiveUser" -Value " "
		}
		$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeInteractiveUser" -Value $LastInteractiveAccess.CreatedDateTime
		$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameInteractiveUser" -Value $LastInteractiveAccess.UserPrincipalName
		$AppData | Add-Member -MemberType NoteProperty -Name "ClientAppUsedInteractiveUser" -Value $LastInteractiveAccess.ClientAppUsed
		
		if($null -ne $LastNonInteractiveAccess) {
			$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value "NonInteractiveUser"
		} else {
			$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeNonInteractiveUser" -Value " "
		}
		$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeNonInteractiveUser" -Value $LastNonInteractiveAccess.CreatedDateTime
		$AppData | Add-Member -MemberType NoteProperty -Name "UserPrincipalNameNonInteractiveUser" -Value $LastNonInteractiveAccess.UserPrincipalName
		$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameNonInteractiveUser" -Value $LastNonInteractiveAccess.ResourceDisplayName
		
		if($null -ne $LastServicePrincipalAccess) {
			$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value "ServicePrincipal"
		} else {
			$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeServicePrincipal" -Value " "
		}
		$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeServicePrincipal" -Value $LastServicePrincipalAccess.CreatedDateTime
		$AppData | Add-Member -MemberType NoteProperty -Name "ServicePrincipalNameServicePrincipal" -Value $LastServicePrincipalAccess.ServicePrincipalName
		$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameServicePrincipal" -Value  $LastServicePrincipalAccess.ResourceDisplayName
		
		if($null -ne $LastManagedIdentityAccess) {
			$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value "ManagedIdentity"
		} else {
			$AppData | Add-Member -MemberType NoteProperty -Name "AccessTypeManagedIdentity" -Value " "
		}
		$AppData | Add-Member -MemberType NoteProperty -Name "LastAccessTimeManagedIdentity" -Value $LastManagedIdentityAccess.CreatedDateTime
		$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityAssociatedResourceId" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.AssociatedResourceId
		$AppData | Add-Member -MemberType NoteProperty -Name "ManagedServiceIdentityMsiType" -Value $LastManagedIdentityAccess.ManagedServiceIdentity.MsiType
		$AppData | Add-Member -MemberType NoteProperty -Name "ResourceDisplayNameManagedIdentity" -Value $LastManagedIdentityAccess.ResourceDisplayName	#>

		$Data += $AppData		
	}
}

if(!([string]::IsNullOrEmpty($Data))) {
	try {
		#Export data in a csv file
		$Data | sort ExpirationDate -Descending | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding Default -Delimiter ";" -ErrorAction Stop
		if($Debug -eq $true) { Write-Output "[DEBUG] : Export csv successfull" }
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

$Style = @"
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body { font-family: ui-sans-serif, -apple-system, system-ui, Segoe UI, Helvetica, Apple Color Emoji, Arial, sans-serif, Segoe UI Emoji, Segoe UI Symbol; font-size: 14px; margin: 20px; padding: 20px; color: #333; }
        h2 { color: #333333; margin-top: 20px; }
        h3 { color: #333333; margin-top: 20px; }
        table { text-align: center; width: 30%; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0px 2px 5px rgba(0, 0, 0, 0.2); }
        td { padding: 12px; border-bottom: 1px solid #ddd; }
        th { background-color: #f0f0f0; padding: 12px; text-align: center; }
    </style>
"@

$BodyContent = @"
Hello,<br /><br />
This email is a reminder regarding Entra ID applications with either a certificate or a client secret expiring soon.<br />
You're receiving this email as you've been identified as a main contact for these applications.<br />
If application(s) list below <b>is still in use</b>, please kindly proceed with its renewal.<br /><br />
"@

#Format html file and apply css style
$html = ($Data | sort ExpirationDate -Descending | select AppDisplayName,AppId,KeyType,ExpirationDate,DaysUntilExpiration | ConvertTo-Html) -replace("<head>",$Style) -replace("<body>","<body>$BodyContent")
if($Debug -eq $true) { "html : {0}" -f $html }
#Display format to avoid content to be seen as array
$html = $html -join "`r`n"	

if(Test-Path $ReportPath) {
	# Email parameters
	if($GlobalReport -eq $true) {
		$MessageAttachement = [Convert]::ToBase64String([IO.File]::ReadAllBytes($ReportPath))
		$SendMailParams = @{
			Message = @{
				Subject = "$Customer - GLOBAL Entra ID Apps Reporting"
				Body = @{
					ContentType = "html"
					Content = "Please find the corresponding report attached."
				}
				ToRecipients = @(							
					@{
						EmailAddress = @{
							Address = $Recipient			 
						}
					}
				)
				Attachments = @(
					@{
						"@odata.type" = "#microsoft.graph.fileAttachment"
						Name = $ReportFile
						ContentType = "text/plain"
						ContentBytes = $MessageAttachement
					}
				)
			}
		SaveToSentItems = "false"
	    }
	} else {	
		$MessageAttachement = [Convert]::ToBase64String([IO.File]::ReadAllBytes($ReportPath))
		$SendMailParams = @{
			Message = @{
				Subject = "$Customer - Entra ID Apps Expiration Reporting"
				Body = @{
					ContentType = "html"
					Content = $html
				}
				ToRecipients = @(
					@{
						EmailAddress = @{
							Address = $Recipient
						}
					}
				)
				Attachments = @(
					@{
						"@odata.type" = "#microsoft.graph.fileAttachment"
						Name = $ReportFile
						ContentType = "text/plain"
						ContentBytes = $MessageAttachement
					}
				)
			}
		SaveToSentItems = "false"
		}
	}

	try {
		Send-MgUserMail -UserId $Sender -BodyParameter $SendMailParams -ErrorAction Stop
		"[OK] : Email sent From:{0} - To:{1}" -f $Sender,$Recipient
	} catch {
		"ERROR : Report could not be sent by mail : {0}" -f $_.Exception[0].Message
	}
}