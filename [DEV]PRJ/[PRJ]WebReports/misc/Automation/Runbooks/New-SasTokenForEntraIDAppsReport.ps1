$RGName = Get-AutomationVariable -Name 'RGName'
$StorageAccountName = Get-AutomationVariable -Name 'StorageAccountName'
$StorageAccountContainerPriv = Get-AutomationVariable -Name 'StorageContainerPriv'
$StorageAccountContainerPub = Get-AutomationVariable -Name 'StorageContainerPub'
$AutomationSubscriptionId = Get-AutomationVariable -Name 'AutomationSubscriptionId'
$WebPageName = "EntraIDApplications"

Connect-AzAccount -Identity

#Operate in the same subscription than automation and storage accounts
Set-AzContext -SubscriptionId $AutomationSubscriptionId

#Retrieve storage context
try {
    $ctx = (Get-AzStorageAccount -ResourceGroupName $RGName -Name $StorageAccountName -ErrorAction Stop).Context
    
    if(!([string]::IsNullOrEmpty($ctx))) {
        Write-Output "[OK] : Storage context retrieved"
    } else {
        Write-Output "[ERROR] : Fail to retrieve storage context : $($_.Exception[0].Message)"
        break
    }
} catch {
    Write-Output "[ERROR] : Fail to generate SAS token : $($_.Exception[0].Message)"
    break
}

#Run once to set the correct CORS rules (Azure Storage blocks cross-origin requests unless CORS rules are explicitly added)
<#try {
    #Define CORS rules
    Set-AzStorageCORSRule -ServiceType Blob `
    -CorsRules @(
        @{
            AllowedHeaders    = "*"
            AllowedOrigins    = "https://dscshow.z28.web.core.windows.net/"
            AllowedMethods    = "GET"
            ExposedHeaders    = "*"
            MaxAgeInSeconds   = 3600
        }
    ) `
    -Context $ctx -ErrorAction Stop
    Write-Output "[OK] : CORS rules added"
} catch {
    Write-Output "[ERROR] : Fail to add CORS rules : $($_.Exception[0].Message)"
    break
} 
#>

#r = read, l = list
$expiry = (Get-Date).AddDays(1)

try {
    $sasToken = New-AzStorageContainerSASToken -Context $ctx -Container $StorageAccountContainerPriv -Permission 'rl' -ExpiryTime $expiry -FullUri:$false -ErrorAction Stop
    Write-Output "[OK] : SAS token generated"
} catch {
    Write-Output "[ERROR] : Fail to generate SAS token : $($_.Exception[0].Message)"
    break
}

#Paths
$templateBlob = "Templates/$WebPageName`.template.html"
$finalBlob = "Main/$WebPageName`.html"
$localTemplate = Join-Path $env:TEMP "Templates\$WebPageName`.template.html"
$localOutput   = Join-Path $env:TEMP "Main\$WebPageName`.html"

#Make sure local folders exist before writing
try { 
    New-Item -ItemType Directory -Path (Split-Path $localTemplate) -Force | Out-Null
    New-Item -ItemType Directory -Path (Split-Path $localOutput) -Force | Out-Null
    Write-Output "[OK] : Temporary directories created"
} catch {
    Write-Output "[ERROR] : Fail to create directories : $($_.Exception[0].Message)"
    break
}

#Download the HTML template from public container
try {
    Get-AzStorageBlobContent -Context $ctx -Container $StorageAccountContainerPub -Blob $templateBlob -Destination $localTemplate -Force -ErrorAction Stop
    Write-Output "[OK] : HTML template downloaded from public container"
} catch {
    Write-Output "[ERROR] : Fail to download the HTML template from public container : $($_.Exception[0].Message)"
    break
}

#Replace the <sas> placeholder with the real token
#Without that question-mark the browser thinks the SAS string is part of the file-path
$sasToken = "?" + $sasToken

try {
    (Get-Content $localTemplate -Raw) -replace "<sas>", $sasToken | Set-Content $localOutput -Encoding UTF8 -ErrorAction Stop
    Write-Output "[DEBUG] Replaced content:"
    Get-Content $localOutput | ForEach-Object { Write-Output $_ }
    Write-Output "[OK] : SAS placeholder replaced with the real token"
} catch {
    Write-Output "[ERROR] : Fail to replace the SAS placerholder with the real token : $($_.Exception[0].Message)" 
    break
}

#Upload the updated file to the public container
try {
    Set-AzStorageBlobContent -Context $ctx -Container $StorageAccountContainerPub -File $localOutput -Blob $finalBlob -Force  -Properties @{"ContentType" = "text/html"} -ErrorAction Stop
    Write-Output "[OK] : Upload with the updated file to the public container successfull"
} catch {
    Write-Output "[ERROR] : Fail to upload the updated file to the public container : $($_.Exception[0].Message)"
}