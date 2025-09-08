[CmdletBinding()]
param(
	[Parameter(Mandatory=$false)]
	[array]$ServiceRequested,
	[Parameter(Mandatory=$false)]
	[array]$ComponentRequested,
	[Parameter(Mandatory=$false)]
	[switch]$FullBackup=$false,
	[Parameter(Mandatory=$false)]
	[switch]$Help=$false
)

<#Parameters :
-Start script to backup all services with all their components in one file : specify -FullBackup
If full backup is requested, no specific service or component can be specified

-Start script to backup one or several service(s) with all their components : specify -ServiceRequested Service1,Service2

-Start script to backup one or several components of one or several service(s) : specify -ComponentRequested Component1,Component2
#>

$FormattedDateForDirectoryFormat = Get-Date -Format "yyyyMMdd"
$FormattedDate = $(Get-Date).ToString("dd_MM_yyyy_HH-mm-ss")

#Script location retrieved automatically
$ScriptLocation = $PSScriptRoot
#If no script location can be retrieved automatically, script location will be current location
if(!($PSScriptRoot -match "\\")) { $ScriptLocation = Get-Location }

#Script name retrieved automatically
$ScriptName = ($MyInvocation.MyCommand.Name).split(".")[0]
#If no script name can be retrieved automatically, script name will be defined as below
if(!($MyInvocation.MyCommand.Name -match "\w")) { $ScriptName = "Export-DSCConfiguration" } 

#Configuration files directory
$DSCExport = $ScriptLocation

#Email settings
$Customer = "GIFI"
$Sender = "alertes@gifi.fr"
$Recipient = "benjamin.poulain@talan.com"
#Connexion
$TenantId = "gifi.onmicrosoft.com"
$ApplicationId = "733fa52d-7d69-4b40-82af-8b55ba6de454"
#Retrieve certificate stored under local machine certificate store
$cert = Get-ChildItem Cert:\LocalMachine\My\ | ?{$_.Subject.StartsWith("CN=M365DSC")}

Set-Location $DSCExport

$Services = @("AzureAD","Exchange","Intune","SharePoint","Teams","OneDrive")
<#All services
$Services = @("AzureAD","Exchange","Intune","SharePoint","Teams","OneDrive","Office365","SecurityCompliance","Planner","PowerPlatform") #>
$AllComponents = @()
$Components = @()

#Display help about script execution
switch($Help) {
	$true {
		Clear-Host
		$Timer = 15
		
		Write-Host "[INFO]: This script can run three parameters that" -ForeGroundColor Yellow -NoNewLine;Write-Host " cannot " -ForeGroundColor Red -NoNewLine;Write-Host "be used at the same time" -ForeGroundColor Yellow
		Write-Host "[INFO]: To perform a full backup : .\PathToScript\Export-DSCConfiguration.ps1" -ForeGroundColor Gray -NoNewLine;Write-Host " -FullBackup" -ForeGroundColor Yellow
		Write-Host "[INFO]: To back up specific services and all of their components : .\PathToScript\Export-DSCConfiguration.ps1" -ForeGroundColor Gray -NoNewLine;Write-Host " -ServiceRequested Service1,Service2" -ForeGroundColor Yellow
		Write-Host "[INFO]: To back up specific components : .\PathToScript\Export-DSCConfiguration.ps1" -ForeGroundColor Gray -NoNewLine;Write-Host " -ComponentRequested Component1,Component2" -ForeGroundColor Yellow
		Write-Host "[INFO]: Services backup currently supported : " -ForeGroundColor Gray -NoNewLine;Write-Host "AzureAD,Exchange,Intune,SharePoint,Teams,OneDrive" -ForeGroundColor Yellow
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
		Write-Host "`tooo Encountered errors" -ForeGroundColor Gray
		
		Start-Sleep -Seconds $Timer
		
		Export-M365DSCConfiguration -LaunchWebUI | Out-Null
		""
		
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
		if($ServiceRequested -or $ComponentRequested) {
			Write-Host "[ERROR]: Full backup requested - you're not allowed to specify any specific service or component" -ForeGroundColor Yellow
			""
			exit
		}
	}
	default { 
		$FullBackup = $false 
		<#Full backup not requested, every service will be backuped and split into different directories
		Unless a specific service is specified #>
		
		if($ServiceRequested) {
			#If specific services are specified, only these services and their components will be backuped
			$Services = $ServiceRequested
			$BackupType = "Services backup"
			Write-Host "[INFO]: Specific services '$($ServiceRequested -join ',')' requested" -ForeGroundColor Yellow
			#No component should be specified when specifying services
			if($ComponentRequested) {
				Write-Host "[ERROR]: Service backup requested - you're not allowed to specify any specific component" -ForeGroundColor Yellow
				exit
			}
		} else {
			if(!($ComponentRequested)) {
				#If no specific services are specified, backup all services and their components
				$BackupType = "Services backup"
				Write-Host "[INFO]: No specific services requested. Backup '$($Services -join ',')'" -ForeGroundColor Yellow
			}
		}
		
		if($ComponentRequested) {
			#If specific components are specified, only these components will be backuped
			$Components = $ComponentRequested
			$BackupType = "Components backup"
			Write-Host "[INFO]: Specific components '$($ComponentRequested -join ",")' requested" -ForeGroundColor Yellow
			if($ServiceRequested) {
				Write-Host "[ERROR]: Component backup requested - you're not allowed to specify any specific service" -ForeGroundColor Yellow
				exit
			}
		}
	}
}

if(!($ComponentRequested)) {
	Foreach($Service in $Services -split ",") {
		Switch($Service) {
			"AzureAD" { 
				if($FullBackup -eq $true) { 
					$Components += @("AADAdministrativeUnit", "AADApplication", "AADAttributeSet", "AADAuthenticationContextClassReference", "AADAuthenticationMethodPolicy", "AADAuthenticationMethodPolicyAuthenticator", "AADAuthenticationMethodPolicyEmail", "AADAuthenticationMethodPolicyFido2", "AADAuthenticationMethodPolicySms", "AADAuthenticationMethodPolicySoftware", "AADAuthenticationMethodPolicyTemporary", "AADAuthenticationMethodPolicyVoice", "AADAuthenticationMethodPolicyX509", "AADAuthenticationStrengthPolicy", "AADAuthorizationPolicy", "AADConditionalAccessPolicy", "AADCrossTenantAccessPolicy", "AADCrossTenantAccessPolicyConfigurationDefault", "AADCrossTenantAccessPolicyConfigurationPartner", "AADEntitlementManagementAccessPackage", "AADEntitlementManagementAccessPackageAssignmentPolicy", "AADEntitlementManagementAccessPackageCatalog", "AADEntitlementManagementAccessPackageCatalogResource", "AADEntitlementManagementConnectedOrganization", "AADExternalIdentityPolicy", "AADGroup", "AADGroupLifecyclePolicy", "AADGroupsNamingPolicy", "AADGroupsSettings", "AADNamedLocationPolicy", "AADRoleDefinition", "AADRoleEligibilityScheduleRequest", "AADRoleSetting", "AADSecurityDefaults", "AADServicePrincipal", "AADSocialIdentityProvider", "AADTenantDetails", "AADTokenLifetimePolicy", "AADUser")
				} else {
					$objAAD = New-Object System.Object
					$objAAD | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("AADAdministrativeUnit", "AADApplication", "AADAttributeSet", "AADAuthenticationContextClassReference", "AADAuthenticationMethodPolicy", "AADAuthenticationMethodPolicyAuthenticator", "AADAuthenticationMethodPolicyEmail", "AADAuthenticationMethodPolicyFido2", "AADAuthenticationMethodPolicySms", "AADAuthenticationMethodPolicySoftware", "AADAuthenticationMethodPolicyTemporary", "AADAuthenticationMethodPolicyVoice", "AADAuthenticationMethodPolicyX509", "AADAuthenticationStrengthPolicy", "AADAuthorizationPolicy", "AADConditionalAccessPolicy", "AADCrossTenantAccessPolicy", "AADCrossTenantAccessPolicyConfigurationDefault", "AADCrossTenantAccessPolicyConfigurationPartner", "AADEntitlementManagementAccessPackage", "AADEntitlementManagementAccessPackageAssignmentPolicy", "AADEntitlementManagementAccessPackageCatalog", "AADEntitlementManagementAccessPackageCatalogResource", "AADEntitlementManagementConnectedOrganization", "AADExternalIdentityPolicy", "AADGroup", "AADGroupLifecyclePolicy", "AADGroupsNamingPolicy", "AADGroupsSettings", "AADNamedLocationPolicy", "AADRoleDefinition", "AADRoleEligibilityScheduleRequest", "AADRoleSetting", "AADSecurityDefaults", "AADServicePrincipal", "AADSocialIdentityProvider", "AADTenantDetails", "AADTokenLifetimePolicy", "AADUser")
					$objAAD | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objAAD
				}
			}
			"Exchange" { 
				if($FullBackup -eq $true) { 
					$Components += @("EXOAcceptedDomain", "EXOActiveSyncDeviceAccessRule", "EXOAddressBookPolicy", "EXOAddressList", "EXOAntiPhishPolicy", "EXOAntiPhishRule", "EXOApplicationAccessPolicy", "EXOAtpPolicyForO365", "EXOAuthenticationPolicy", "EXOAuthenticationPolicyAssignment", "EXOAvailabilityAddressSpace", "EXOAvailabilityConfig", "EXOCalendarProcessing", "EXOCASMailboxPlan", "EXOCASMailboxSettings", "EXOClientAccessRule", "EXODataClassification", "EXODataEncryptionPolicy", "EXODistributionGroup", "EXODkimSigningConfig", "EXOEmailAddressPolicy", "EXOGlobalAddressList", "EXOGroupSettings", "EXOHostedConnectionFilterPolicy", "EXOHostedContentFilterPolicy", "EXOHostedContentFilterRule", "EXOHostedOutboundSpamFilterPolicy", "EXOHostedOutboundSpamFilterRule", "EXOInboundConnector", "EXOIntraOrganizationConnector", "EXOIRMConfiguration", "EXOJournalRule", "EXOMailboxAutoReplyConfiguration", "EXOMailboxCalendarFolder", "EXOMailboxPermission", "EXOMailboxPlan", "EXOMailboxSettings", "EXOMailContact", "EXOMailTips", "EXOMalwareFilterPolicy", "EXOMalwareFilterRule", "EXOManagementRole", "EXOManagementRoleAssignment", "EXOMessageClassification", "EXOMobileDeviceMailboxPolicy", "EXOOfflineAddressBook", "EXOOMEConfiguration", "EXOOnPremisesOrganization", "EXOOrganizationConfig", "EXOOrganizationRelationship", "EXOOutboundConnector", "EXOOwaMailboxPolicy", "EXOPartnerApplication", "EXOPerimeterConfiguration", "EXOPlace", "EXOPolicyTipConfig", "EXOQuarantinePolicy", "EXORecipientPermission", "EXORemoteDomain", "EXOReportSubmissionPolicy", "EXOReportSubmissionRule", "EXOResourceConfiguration", "EXORoleAssignmentPolicy", "EXORoleGroup", "EXOSafeAttachmentPolicy", "EXOSafeAttachmentRule", "EXOSafeLinksPolicy", "EXOSafeLinksRule", "EXOSharedMailbox", "EXOSharingPolicy", "EXOTransportConfig", "EXOTransportRule")
				} else {
					$objEXC = New-Object System.Object
					$objEXC | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("EXOAcceptedDomain", "EXOActiveSyncDeviceAccessRule", "EXOAddressBookPolicy", "EXOAddressList", "EXOAntiPhishPolicy", "EXOAntiPhishRule", "EXOApplicationAccessPolicy", "EXOAtpPolicyForO365", "EXOAuthenticationPolicy", "EXOAuthenticationPolicyAssignment", "EXOAvailabilityAddressSpace", "EXOAvailabilityConfig", "EXOCalendarProcessing", "EXOCASMailboxPlan", "EXOCASMailboxSettings", "EXOClientAccessRule", "EXODataClassification", "EXODataEncryptionPolicy", "EXODistributionGroup", "EXODkimSigningConfig", "EXOEmailAddressPolicy", "EXOGlobalAddressList", "EXOGroupSettings", "EXOHostedConnectionFilterPolicy", "EXOHostedContentFilterPolicy", "EXOHostedContentFilterRule", "EXOHostedOutboundSpamFilterPolicy", "EXOHostedOutboundSpamFilterRule", "EXOInboundConnector", "EXOIntraOrganizationConnector", "EXOIRMConfiguration", "EXOJournalRule", "EXOMailboxAutoReplyConfiguration", "EXOMailboxCalendarFolder", "EXOMailboxPermission", "EXOMailboxPlan", "EXOMailboxSettings", "EXOMailContact", "EXOMailTips", "EXOMalwareFilterPolicy", "EXOMalwareFilterRule", "EXOManagementRole", "EXOManagementRoleAssignment", "EXOMessageClassification", "EXOMobileDeviceMailboxPolicy", "EXOOfflineAddressBook", "EXOOMEConfiguration", "EXOOnPremisesOrganization", "EXOOrganizationConfig", "EXOOrganizationRelationship", "EXOOutboundConnector", "EXOOwaMailboxPolicy", "EXOPartnerApplication", "EXOPerimeterConfiguration", "EXOPlace", "EXOPolicyTipConfig", "EXOQuarantinePolicy", "EXORecipientPermission", "EXORemoteDomain", "EXOReportSubmissionPolicy", "EXOReportSubmissionRule", "EXOResourceConfiguration", "EXORoleAssignmentPolicy", "EXORoleGroup", "EXOSafeAttachmentPolicy", "EXOSafeAttachmentRule", "EXOSafeLinksPolicy", "EXOSafeLinksRule", "EXOSharedMailbox", "EXOSharingPolicy", "EXOTransportConfig", "EXOTransportRule")
					$objEXC | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objEXC
				}
			}
			"Intune" { 
				if($FullBackup -eq $true) { 
					$Components += @("IntuneAccountProtectionLocalAdministratorPasswordSolutionPolicy", "IntuneAccountProtectionLocalUserGroupMembershipPolicy", "IntuneAccountProtectionPolicy", "IntuneAntivirusPolicyWindows10SettingCatalog", "IntuneAppConfigurationPolicy", "IntuneApplicationControlPolicyWindows10", "IntuneAppProtectionPolicyAndroid", "IntuneAppProtectionPolicyiOS", "IntuneASRRulesPolicyWindows10", "IntuneAttackSurfaceReductionRulesPolicyWindows10ConfigManager", "IntuneDeviceAndAppManagementAssignmentFilter", "IntuneDeviceCategory", "IntuneDeviceCleanupRule", "IntuneDeviceCompliancePolicyAndroid", "IntuneDeviceCompliancePolicyAndroidDeviceOwner", "IntuneDeviceCompliancePolicyAndroidWorkProfile", "IntuneDeviceCompliancePolicyiOs", "IntuneDeviceCompliancePolicyMacOS", "IntuneDeviceCompliancePolicyWindows10", "IntuneDeviceConfigurationAdministrativeTemplatePolicyWindows10", "IntuneDeviceConfigurationCustomPolicyWindows10", "IntuneDeviceConfigurationDefenderForEndpointOnboardingPolicyWindows10", "IntuneDeviceConfigurationDeliveryOptimizationPolicyWindows10", "IntuneDeviceConfigurationDomainJoinPolicyWindows10", "IntuneDeviceConfigurationEmailProfilePolicyWindows10", "IntuneDeviceConfigurationEndpointProtectionPolicyWindows10", "IntuneDeviceConfigurationFirmwareInterfacePolicyWindows10", "IntuneDeviceConfigurationHealthMonitoringConfigurationPolicyWindows10", "IntuneDeviceConfigurationIdentityProtectionPolicyWindows10", "IntuneDeviceConfigurationImportedPfxCertificatePolicyWindows10", "IntuneDeviceConfigurationKioskPolicyWindows10", "IntuneDeviceConfigurationNetworkBoundaryPolicyWindows10", "IntuneDeviceConfigurationPkcsCertificatePolicyWindows10", "IntuneDeviceConfigurationPolicyAndroidDeviceAdministrator", "IntuneDeviceConfigurationPolicyAndroidDeviceOwner", "IntuneDeviceConfigurationPolicyAndroidOpenSourceProject", "IntuneDeviceConfigurationPolicyAndroidWorkProfile", "IntuneDeviceConfigurationPolicyiOS", "IntuneDeviceConfigurationPolicyMacOS", "IntuneDeviceConfigurationPolicyWindows10", "IntuneDeviceConfigurationSCEPCertificatePolicyWindows10", "IntuneDeviceConfigurationSecureAssessmentPolicyWindows10", "IntuneDeviceConfigurationSharedMultiDevicePolicyWindows10", "IntuneDeviceConfigurationTrustedCertificatePolicyWindows10", "IntuneDeviceConfigurationVpnPolicyWindows10", "IntuneDeviceConfigurationWindowsTeamPolicyWindows10", "IntuneDeviceConfigurationWiredNetworkPolicyWindows10", "IntuneDeviceEnrollmentLimitRestriction", "IntuneDeviceEnrollmentPlatformRestriction", "IntuneDeviceEnrollmentStatusPageWindows10", "IntuneEndpointDetectionAndResponsePolicyWindows10", "IntuneExploitProtectionPolicyWindows10SettingCatalog", "IntunePolicySets", "IntuneRoleAssignment", "IntuneRoleDefinition", "IntuneSettingCatalogASRRulesPolicyWindows10", "IntuneSettingCatalogCustomPolicyWindows10", "IntuneWiFiConfigurationPolicyAndroidDeviceAdministrator", "IntuneWifiConfigurationPolicyAndroidEnterpriseDeviceOwner", "IntuneWifiConfigurationPolicyAndroidEnterpriseWorkProfile", "IntuneWifiConfigurationPolicyAndroidForWork", "IntuneWifiConfigurationPolicyAndroidOpenSourceProject", "IntuneWifiConfigurationPolicyIOS", "IntuneWifiConfigurationPolicyMacOS", "IntuneWifiConfigurationPolicyWindows10", "IntuneWindowsAutopilotDeploymentProfileAzureADHybridJoined", "IntuneWindowsAutopilotDeploymentProfileAzureADJoined", "IntuneWindowsInformationProtectionPolicyWindows10MdmEnrolled", "IntuneWindowsUpdateForBusinessFeatureUpdateProfileWindows10", "IntuneWindowsUpdateForBusinessRingUpdateProfileWindows10")
				} else {
					$objINT = New-Object System.Object
					$objINT | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("IntuneAccountProtectionLocalAdministratorPasswordSolutionPolicy", "IntuneAccountProtectionLocalUserGroupMembershipPolicy", "IntuneAccountProtectionPolicy", "IntuneAntivirusPolicyWindows10SettingCatalog", "IntuneAppConfigurationPolicy", "IntuneApplicationControlPolicyWindows10", "IntuneAppProtectionPolicyAndroid", "IntuneAppProtectionPolicyiOS", "IntuneASRRulesPolicyWindows10", "IntuneAttackSurfaceReductionRulesPolicyWindows10ConfigManager", "IntuneDeviceAndAppManagementAssignmentFilter", "IntuneDeviceCategory", "IntuneDeviceCleanupRule", "IntuneDeviceCompliancePolicyAndroid", "IntuneDeviceCompliancePolicyAndroidDeviceOwner", "IntuneDeviceCompliancePolicyAndroidWorkProfile", "IntuneDeviceCompliancePolicyiOs", "IntuneDeviceCompliancePolicyMacOS", "IntuneDeviceCompliancePolicyWindows10", "IntuneDeviceConfigurationAdministrativeTemplatePolicyWindows10", "IntuneDeviceConfigurationCustomPolicyWindows10", "IntuneDeviceConfigurationDefenderForEndpointOnboardingPolicyWindows10", "IntuneDeviceConfigurationDeliveryOptimizationPolicyWindows10", "IntuneDeviceConfigurationDomainJoinPolicyWindows10", "IntuneDeviceConfigurationEmailProfilePolicyWindows10", "IntuneDeviceConfigurationEndpointProtectionPolicyWindows10", "IntuneDeviceConfigurationFirmwareInterfacePolicyWindows10", "IntuneDeviceConfigurationHealthMonitoringConfigurationPolicyWindows10", "IntuneDeviceConfigurationIdentityProtectionPolicyWindows10", "IntuneDeviceConfigurationImportedPfxCertificatePolicyWindows10", "IntuneDeviceConfigurationKioskPolicyWindows10", "IntuneDeviceConfigurationNetworkBoundaryPolicyWindows10", "IntuneDeviceConfigurationPkcsCertificatePolicyWindows10", "IntuneDeviceConfigurationPolicyAndroidDeviceAdministrator", "IntuneDeviceConfigurationPolicyAndroidDeviceOwner", "IntuneDeviceConfigurationPolicyAndroidOpenSourceProject", "IntuneDeviceConfigurationPolicyAndroidWorkProfile", "IntuneDeviceConfigurationPolicyiOS", "IntuneDeviceConfigurationPolicyMacOS", "IntuneDeviceConfigurationPolicyWindows10", "IntuneDeviceConfigurationSCEPCertificatePolicyWindows10", "IntuneDeviceConfigurationSecureAssessmentPolicyWindows10", "IntuneDeviceConfigurationSharedMultiDevicePolicyWindows10", "IntuneDeviceConfigurationTrustedCertificatePolicyWindows10", "IntuneDeviceConfigurationVpnPolicyWindows10", "IntuneDeviceConfigurationWindowsTeamPolicyWindows10", "IntuneDeviceConfigurationWiredNetworkPolicyWindows10", "IntuneDeviceEnrollmentLimitRestriction", "IntuneDeviceEnrollmentPlatformRestriction", "IntuneDeviceEnrollmentStatusPageWindows10", "IntuneEndpointDetectionAndResponsePolicyWindows10", "IntuneExploitProtectionPolicyWindows10SettingCatalog", "IntunePolicySets", "IntuneRoleAssignment", "IntuneRoleDefinition", "IntuneSettingCatalogASRRulesPolicyWindows10", "IntuneSettingCatalogCustomPolicyWindows10", "IntuneWiFiConfigurationPolicyAndroidDeviceAdministrator", "IntuneWifiConfigurationPolicyAndroidEnterpriseDeviceOwner", "IntuneWifiConfigurationPolicyAndroidEnterpriseWorkProfile", "IntuneWifiConfigurationPolicyAndroidForWork", "IntuneWifiConfigurationPolicyAndroidOpenSourceProject", "IntuneWifiConfigurationPolicyIOS", "IntuneWifiConfigurationPolicyMacOS", "IntuneWifiConfigurationPolicyWindows10", "IntuneWindowsAutopilotDeploymentProfileAzureADHybridJoined", "IntuneWindowsAutopilotDeploymentProfileAzureADJoined", "IntuneWindowsInformationProtectionPolicyWindows10MdmEnrolled", "IntuneWindowsUpdateForBusinessFeatureUpdateProfileWindows10", "IntuneWindowsUpdateForBusinessRingUpdateProfileWindows10")
					$objINT | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objINT
				}	
			}
			"OneDrive" {
				if($FullBackup -eq $true) { 
					$Components += @("ODSettings")
				} else {
					$objOD = New-Object System.Object
					$objOD | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("ODSettings")
					$objOD | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objOD
				}
			}
			"SharePoint" {
				if($FullBackup -eq $true) { 
					$Components += @("SPOAccessControlSettings", "SPOBrowserIdleSignout", "SPOHomeSite", "SPOHubSite", "SPOOrgAssetsLibrary", "SPOPropertyBag", "SPOSearchManagedProperty", "SPOSearchResultSource", "SPOSharingSettings", "SPOSite", "SPOSiteAuditSettings", "SPOSiteDesign", "SPOSiteDesignRights", "SPOSiteGroup", "SPOSiteScript", "SPOStorageEntity", "SPOTenantCdnEnabled", "SPOTenantCdnPolicy", "SPOTenantSettings", "SPOTheme", "SPOUserProfileProperty", "SPOApp")
				} else {
					$objSP = New-Object System.Object
					$objSP | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("SPOAccessControlSettings", "SPOBrowserIdleSignout", "SPOHomeSite", "SPOHubSite", "SPOOrgAssetsLibrary", "SPOPropertyBag", "SPOSearchManagedProperty", "SPOSearchResultSource", "SPOSharingSettings", "SPOSite", "SPOSiteAuditSettings", "SPOSiteDesign", "SPOSiteDesignRights", "SPOSiteGroup", "SPOSiteScript", "SPOStorageEntity", "SPOTenantCdnEnabled", "SPOTenantCdnPolicy", "SPOTenantSettings", "SPOTheme", "SPOUserProfileProperty", "SPOApp")
					$objSP | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objSP
				}
			}
			"Teams" {
				if($FullBackup -eq $true) { 
					$Components += @("TeamsAppPermissionPolicy", "TeamsAppSetupPolicy", "TeamsAudioConferencingPolicy", "TeamsCallHoldPolicy", "TeamsCallingPolicy", "TeamsCallParkPolicy", "TeamsCallQueue", "TeamsChannel", "TeamsChannelsPolicy", "TeamsChannelTab", "TeamsClientConfiguration", "TeamsComplianceRecordingPolicy", "TeamsCortanaPolicy", "TeamsDialInConferencingTenantSettings", "TeamsEmergencyCallingPolicy", "TeamsEmergencyCallRoutingPolicy", "TeamsEnhancedEncryptionPolicy", "TeamsEventsPolicy", "TeamsFederationConfiguration", "TeamsFeedbackPolicy", "TeamsFilesPolicy", "TeamsGroupPolicyAssignment", "TeamsGuestCallingConfiguration", "TeamsGuestMeetingConfiguration", "TeamsGuestMessagingConfiguration", "TeamsIPPhonePolicy", "TeamsMeetingBroadcastConfiguration", "TeamsMeetingBroadcastPolicy", "TeamsMeetingConfiguration", "TeamsMeetingPolicy", "TeamsMessagingPolicy", "TeamsMobilityPolicy", "TeamsNetworkRoamingPolicy", "TeamsOnlineVoicemailPolicy", "TeamsOnlineVoicemailUserSettings", "TeamsOnlineVoiceUser", "TeamsOrgWideAppSettings", "TeamsPstnUsage", "TeamsShiftsPolicy", "TeamsTeam", "TeamsTemplatesPolicy", "TeamsTenantDialPlan", "TeamsTenantNetworkRegion", "TeamsTenantNetworkSite", "TeamsTenantNetworkSubnet", "TeamsTenantTrustedIPAddress", "TeamsTranslationRule", "TeamsUnassignedNumberTreatment", "TeamsUpdateManagementPolicy", "TeamsUpgradeConfiguration", "TeamsUpgradePolicy", "TeamsUser", "TeamsUserCallingSettings", "TeamsUserPolicyAssignment", "TeamsVdiPolicy", "TeamsVoiceRoute", "TeamsVoiceRoutingPolicy", "TeamsWorkloadPolicy")
				} else {
					$objTMS = New-Object System.Object
					$objTMS | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("TeamsAppPermissionPolicy", "TeamsAppSetupPolicy", "TeamsAudioConferencingPolicy", "TeamsCallHoldPolicy", "TeamsCallingPolicy", "TeamsCallParkPolicy", "TeamsCallQueue", 	"TeamsChannel", "TeamsChannelsPolicy", "TeamsChannelTab", "TeamsClientConfiguration", "TeamsComplianceRecordingPolicy", "TeamsCortanaPolicy", "TeamsDialInConferencingTenantSettings", "TeamsEmergencyCallingPolicy", "TeamsEmergencyCallRoutingPolicy", "TeamsEnhancedEncryptionPolicy", "TeamsEventsPolicy", "TeamsFederationConfiguration", "TeamsFeedbackPolicy", "TeamsFilesPolicy", "TeamsGroupPolicyAssignment", "TeamsGuestCallingConfiguration", "TeamsGuestMeetingConfiguration", "TeamsGuestMessagingConfiguration", "TeamsIPPhonePolicy", "TeamsMeetingBroadcastConfiguration", "TeamsMeetingBroadcastPolicy", "TeamsMeetingConfiguration", "TeamsMeetingPolicy", "TeamsMessagingPolicy", "TeamsMobilityPolicy", "TeamsNetworkRoamingPolicy", "TeamsOnlineVoicemailPolicy", "TeamsOnlineVoicemailUserSettings", "TeamsOnlineVoiceUser", "TeamsOrgWideAppSettings", "TeamsPstnUsage", "TeamsShiftsPolicy", "TeamsTeam", "TeamsTemplatesPolicy", "TeamsTenantDialPlan", "TeamsTenantNetworkRegion", "TeamsTenantNetworkSite", "TeamsTenantNetworkSubnet", "TeamsTenantTrustedIPAddress", "TeamsTranslationRule", "TeamsUnassignedNumberTreatment", "TeamsUpdateManagementPolicy", "TeamsUpgradeConfiguration", "TeamsUpgradePolicy", "TeamsUser", "TeamsUserCallingSettings", "TeamsUserPolicyAssignment", "TeamsVdiPolicy", "TeamsVoiceRoute", "TeamsVoiceRoutingPolicy", "TeamsWorkloadPolicy")
					$objTMS | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objTMS
				}
			}
			<#"Office365" {
				if($FullBackup -eq $true) { 
					$Components += @("O365AdminAuditLogConfig", "O365Group", "O365OrgCustomizationSetting", "O365OrgSettings", "O365SearchAndIntelligenceConfigurations")
				} else {
					$obj365 = New-Object System.Object
					$obj365 | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("O365AdminAuditLogConfig", "O365Group", "O365OrgCustomizationSetting", "O365OrgSettings", "O365SearchAndIntelligenceConfigurations")
					$obj365 | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $obj365
				}
			} #>
			<#"SecurityCompliance" {
				if($FullBackup -eq $true) { 
					$Components += @("SCAuditConfigurationPolicy", "SCAutoSensitivityLabelPolicy", "SCAutoSensitivityLabelRule", "SCCaseHoldPolicy", "SCCaseHoldRule", "SCComplianceCase", "SCComplianceSearch", "SCComplianceSearchAction", "SCComplianceTag", "SCDeviceConditionalAccessPolicy", "SCDeviceConfigurationPolicy", "SCDLPCompliancePolicy", "SCDLPComplianceRule", "SCFilePlanPropertyAuthority", "SCFilePlanPropertyCategory", "SCFilePlanPropertyCitation", "SCFilePlanPropertyDepartment", "SCFilePlanPropertyReferenceId", "SCFilePlanPropertySubCategory", "SCLabelPolicy", "SCProtectionAlert", "SCRetentionCompliancePolicy", "SCRetentionComplianceRule", "SCRetentionEventType", "SCSecurityFilter", "SCSensitivityLabel", "SCSupervisoryReviewPolicy", "SCSupervisoryReviewRule")
				} else {
					$objSC = New-Object System.Object
					$objSC | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("SCAuditConfigurationPolicy", "SCAutoSensitivityLabelPolicy", "SCAutoSensitivityLabelRule", "SCCaseHoldPolicy", "SCCaseHoldRule", "SCComplianceCase", "SCComplianceSearch", "SCComplianceSearchAction", "SCComplianceTag", "SCDeviceConditionalAccessPolicy", "SCDeviceConfigurationPolicy", "SCDLPCompliancePolicy", "SCDLPComplianceRule", "SCFilePlanPropertyAuthority", "SCFilePlanPropertyCategory", "SCFilePlanPropertyCitation", "SCFilePlanPropertyDepartment", "SCFilePlanPropertyReferenceId", "SCFilePlanPropertySubCategory", "SCLabelPolicy", "SCProtectionAlert", "SCRetentionCompliancePolicy", "SCRetentionComplianceRule", "SCRetentionEventType", "SCSecurityFilter", "SCSensitivityLabel", "SCSupervisoryReviewPolicy", "SCSupervisoryReviewRule")
					$objSC | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objSC
				}
			} #>
			<#"Planner" {
				if($FullBackup -eq $true) { 
					$Components += @("PlannerBucket", "PlannerPlan", "PlannerTask")
				} else {
					$objPLA = New-Object System.Object
					$objPLA | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("PlannerBucket", "PlannerPlan", "PlannerTask")
					$objPLA | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objPLA
				}
			} #>
			<#"PowerPlatform" {
				if($FullBackup -eq $true) { 
					$Components += @("PPPowerAppsEnvironment", "PPTenantIsolationSettings", "PPTenantSettings")
				} else {
					$objPP = New-Object System.Object
					$objPP | Add-Member -MemberType NoteProperty -Name "Service" -Value $Service
					$Components = @("PPPowerAppsEnvironment", "PPTenantIsolationSettings", "PPTenantSettings")
					$objPP | Add-Member -MemberType NoteProperty -Name "Components" -Value $Components
					$AllComponents += $objPP
				}
			} #>
		
			default {  
				Write-Host "[ERROR]: No component found for $Service" -ForeGroundColor Yellow
				break
			}
		}
	}
}

""
Write-Host "[INFO]: Script starting location '$DSCExport'" -ForeGroundColor Yellow

if($FullBackup -eq $true) { 
	Write-Host "[INFO]: Starting full backup for components $($Components -join ',')" -ForeGroundColor Yellow
	#If full backup is specified, back up every service and component in one file
	$TranscriptFile = "$DSCExport\Transcript_FullBackup_$FormattedDate.log"
	
	Start-Transcript -Path $TranscriptFile -Force
	
	Export-M365DSCConfiguration -CertificateThumbprint $cert.Thumbprint -TenantId $TenantId -ApplicationId $ApplicationId -Path "$DSCExport\FullBackup_$FormattedDateForDirectoryFormat" -FileName "FullBackup.ps1" -Components $Components #298 components on 2024/03/28

	Stop-Transcript
	
	#Compile the full backup script into a compiled file used to restore the configuration
	#$DSCExport\FullBackup\FullBackup_$FormattedDate.ps1
	
	$MessageAttachement = [Convert]::ToBase64String([IO.File]::ReadAllBytes($TranscriptFile))
	$Subject = "$Customer - $BackupType completed"
	$TranscriptContent = Get-Content $TranscriptFile
	$ScriptStart = ($TranscriptContent | select -First 1 -Skip 2) -replace ("\D","")
	$YearStart = $ScriptStart[0..3] -join "";$MonthStart = $ScriptStart[4..5] -join "";$DayStart = $ScriptStart[6..7] -join "";$HourStart = $ScriptStart[8..9] -join "";$MinutStart = $ScriptStart[10..11] -join "";$SecondStart = $ScriptStart[12..13] -join ""
	$ScriptStartDate = (Get-Date -Year $YearStart -Month $MonthStart -Day $DayStart -Hour $HourStart -Minute $MinutStart -Second $SecondStart).GetDateTimeFormats()[50]
	$ScriptEnd = (Get-Content $TranscriptFile | select -Last 1 -Skip 1) -replace ("\D","")
	$YearEnd = $ScriptEnd[0..3] -join "";$MonthEnd = $ScriptEnd[4..5] -join "";$DayEnd = $ScriptEnd[6..7] -join "";$HourEnd = $ScriptEnd[8..9] -join "";$MinutEnd = $ScriptEnd[10..11] -join "";$SecondEnd = $ScriptEnd[12..13] -join ""
	$ScriptEndDate = (Get-Date -Year $YearEnd -Month $MonthEnd -Day $DayEnd -Hour $HourEnd -Minute $MinutEnd -Second $SecondEnd).GetDateTimeFormats()[50]
	$BodyContent = "Backup start : $ScriptStartDate <br />Backup end : $ScriptEndDate <br/><br/>"
	
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
	Write-Host "[INFO]: Sending email From:'$Sender' To:'$Recipient' with Subject:'$Subject'" -ForeGroundColor Yellow
	
	try {
		Send-MgUserMail -UserId $Sender -BodyParameter $SendMailParams -ErrorAction Stop
	} catch {
		Write-Host "[ERROR]: Report could not be sent by mail : $($_.Exception[0].Message)" -ForeGroundColor Yellow
	}
	
	
	
} elseif(!($ComponentRequested)) {
	Write-Host "[INFO]: Starting backup for services $($Services -join ',')" -ForeGroundColor Yellow
	
	#If no component is specified, back up by service and components
	Foreach($Service in ($Services -split ",")) {
		#Transcript file
		$TranscriptFile = "$DSCExport\Transcript_$Service`_$FormattedDate.log"

		Start-Transcript -Path $TranscriptFile -Force
		
		Foreach($Component in $(($AllComponents | ?{$_.Service -eq $Service}).Components)) {
			Export-M365DSCConfiguration -CertificateThumbprint $cert.Thumbprint -TenantId $TenantId -ApplicationId $ApplicationId -Path "$DSCExport\$Service`_$FormattedDateForDirectoryFormat" -FileName "$Component.ps1" -Components $Component
		}
		
		Stop-Transcript
		
		$MessageAttachement = [Convert]::ToBase64String([IO.File]::ReadAllBytes($TranscriptFile))
		$Subject = "$Customer - $BackupType for $Service completed"
		$TranscriptContent = Get-Content $TranscriptFile
		$ScriptStart = ($TranscriptContent | select -First 1 -Skip 2) -replace ("\D","")
		$YearStart = $ScriptStart[0..3] -join "";$MonthStart = $ScriptStart[4..5] -join "";$DayStart = $ScriptStart[6..7] -join "";$HourStart = $ScriptStart[8..9] -join "";$MinutStart = $ScriptStart[10..11] -join "";$SecondStart = $ScriptStart[12..13] -join ""
		$ScriptStartDate = (Get-Date -Year $YearStart -Month $MonthStart -Day $DayStart -Hour $HourStart -Minute $MinutStart -Second $SecondStart).GetDateTimeFormats()[50]
		$ScriptEnd = (Get-Content $TranscriptFile | select -Last 1 -Skip 1) -replace ("\D","")
		$YearEnd = $ScriptEnd[0..3] -join "";$MonthEnd = $ScriptEnd[4..5] -join "";$DayEnd = $ScriptEnd[6..7] -join "";$HourEnd = $ScriptEnd[8..9] -join "";$MinutEnd = $ScriptEnd[10..11] -join "";$SecondEnd = $ScriptEnd[12..13] -join ""
		$ScriptEndDate = (Get-Date -Year $YearEnd -Month $MonthEnd -Day $DayEnd -Hour $HourEnd -Minute $MinutEnd -Second $SecondEnd).GetDateTimeFormats()[50]
		$BodyContent = "Backup start : $ScriptStartDate <br />Backup end : $ScriptEndDate <br/><br/>"
		
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
		Write-Host "[INFO]: Sending email From:'$Sender' To:'$Recipient' with Subject:'$Subject'" -ForeGroundColor Yellow
		
		try {
			Send-MgUserMail -UserId $Sender -BodyParameter $SendMailParams -ErrorAction Stop
		} catch {
			Write-Host "[ERROR]: Report could not be sent by mail : $($_.Exception[0].Message)" -ForeGroundColor Yellow
		}
	}
} else {
	#If components are specified they can come from different services
	#Transcript file
	$TranscriptFile = "$DSCExport\Transcript_$($Components -join '')`_$FormattedDate.log"
	
	Start-Transcript -Path $TranscriptFile -Force
	
	#Back up in a different directory with no category as custom components
	Write-Host "[INFO]: Starting backup for components $($Components -join ',')" -ForeGroundColor Yellow
	Export-M365DSCConfiguration -CertificateThumbprint $cert.Thumbprint -TenantId $TenantId -ApplicationId $ApplicationId -Path "$DSCExport\customComponents_$FormattedDateForDirectoryFormat" -FileName "$($Components -join '').ps1" -Components $Components
	
	#Compile the backup script into a compiled file used to restore the configuration
	#$DSCExport\customComponents\EXOTransportRuleAADAuthenticationMethodPolicy_28_03_2024_22-50-30.ps1
	
	Stop-Transcript
	
	$MessageAttachement = [Convert]::ToBase64String([IO.File]::ReadAllBytes($TranscriptFile))
	$Subject = "$Customer - $BackupType for custom components completed"
	$TranscriptContent = Get-Content $TranscriptFile
	$ScriptStart = ($TranscriptContent | select -First 1 -Skip 2) -replace ("\D","")
	$YearStart = $ScriptStart[0..3] -join "";$MonthStart = $ScriptStart[4..5] -join "";$DayStart = $ScriptStart[6..7] -join "";$HourStart = $ScriptStart[8..9] -join "";$MinutStart = $ScriptStart[10..11] -join "";$SecondStart = $ScriptStart[12..13] -join ""
	$ScriptStartDate = (Get-Date -Year $YearStart -Month $MonthStart -Day $DayStart -Hour $HourStart -Minute $MinutStart -Second $SecondStart).GetDateTimeFormats()[50]
	$ScriptEnd = (Get-Content $TranscriptFile | select -Last 1 -Skip 1) -replace ("\D","")
	$YearEnd = $ScriptEnd[0..3] -join "";$MonthEnd = $ScriptEnd[4..5] -join "";$DayEnd = $ScriptEnd[6..7] -join "";$HourEnd = $ScriptEnd[8..9] -join "";$MinutEnd = $ScriptEnd[10..11] -join "";$SecondEnd = $ScriptEnd[12..13] -join ""
	$ScriptEndDate = (Get-Date -Year $YearEnd -Month $MonthEnd -Day $DayEnd -Hour $HourEnd -Minute $MinutEnd -Second $SecondEnd).GetDateTimeFormats()[50]
	$BodyContent = "Backup start : $ScriptStartDate <br />Backup end:$ScriptEndDate <br/><br/>Components extracted : $((($TranscriptContent | Select-String 'Exporting Microsoft 365') -split ': ')[-1]) <br/><br/>"
	
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
	
	Write-Host "[INFO]: Sending email From:'$Sender' To:'$Recipient' with Subject:'$Subject'" -ForeGroundColor Yellow
	
	try {
		Send-MgUserMail -UserId $Sender -BodyParameter $SendMailParams -ErrorAction Stop
	} catch {
		Write-Host "[ERROR]: Report could not be sent by mail : $($_.Exception[0].Message)" -ForeGroundColor Yellow
	}
}

#Remove logs and backups older than 2 months
#Get-ChildItem $DSCExport | ?{($_.PsIsContainer -or $_.Name.EndsWith(".log")) -and ($_.LastWriteTime -lt (Get-Date).AddMonths(-2))} | Remove-Item -Recurse -Force -WhatIf