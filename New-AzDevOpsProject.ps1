[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $TenantId = [string]::Empty,
    [Parameter(Mandatory = $true)]
    [string]
    $SubscriptionId = [string]::Empty,
    [Parameter(Mandatory = $true)]
    [string]
    $OrganizationName = [string]::Empty,
    [Parameter(Mandatory = $true)]
    [string]
    $PAT = [string]::Empty,
    [Parameter(Mandatory = $true)]
    [string]
    $ProjectName = [string]::Empty,
    [Parameter()]
    [string]
    $ProjectDescription = "Project created with Azure DevOps Rest API"
)

function Initialize-Modules {
    param (
        [string[]]$Modules = @("Az", "AzureADPreview") 
    )

    foreach ($module in $Modules) {
        if (-not (Get-InstalledModule -Name "$($module)" -ErrorAction Ignore)) {
            Write-Host "Installing $($module) module..." -ForegroundColor White -BackgroundColor Black
            Install-Module -Name "$($module)" -Force -AllowClobber
            Import-Module "$($module)"
        }
        else {
            Write-Host "Updating $($module) module..." -ForegroundColor White -BackgroundColor Black
            Update-Module -Name "$($module)"
        }
    }
}

function New-AzADAppForDevops {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$OrganizationName,
  
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ApplicationName,
  
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,
  
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SubscriptionId
    )
    $applicationName = "$($OrganizationName)-$ApplicationName"
    $appIdUri = "https://VisualStudio/SPN/$($applicationName)"
    $appReplyUrls = @($appIdUri)
    
    if ($SubscriptionId) {
        Write-Host "Connect AzAccount..." -ForegroundColor White -BackgroundColor Black
        Connect-AzAccount -TenantId $TenantId -SubscriptionId "$($SubscriptionId)" | Out-Null
        Set-AzContext -TenantId $TenantId -SubscriptionId $SubscriptionId | Out-Null
    }
    else {
        throw "INVALID_SUBSCRIPTION_ID"
    }
  
    Write-Host "Checking application $($applicationName)..." -ForegroundColor White -BackgroundColor Black
    $aadApplication = Get-AzADApplication -DisplayName $applicationName
    if ($null -ne $aadApplication) {
        #Application exits, remove it
        Remove-AzADApplication -DisplayName $applicationName -Force
        $aadApplication = $null
    }
    if ($null -eq $aadApplication) {
        #Create application
        Write-Host "Creating application $($applicationName)..." -ForegroundColor White -BackgroundColor Black
        $aadApplication = New-AzADApplication -DisplayName $applicationName -HomePage "https://VisualStudio/SPN" -IdentifierUris $appIdUri -ReplyUrls $appReplyUrls
        Write-Host "Application $($applicationName) is created..." -ForegroundColor White -BackgroundColor Black
        $aadApplication
        if (-not $aadApplication) {
            throw "ADAPPLICATION_ERROR"
        }

        #Create ServicePrincipal and Owner role
        Write-Host "Adding service principal to $($applicationName)..." -ForegroundColor White -BackgroundColor Black
        $servicePrincipal = New-AzADServicePrincipal -DisplayName $applicationName -Role "Owner" -ApplicationId $aadApplication.ApplicationId
        if (-not $servicePrincipal) {
            throw "SERVICEPRINCIPAL_ERROR"
        }
  
        $servicePrincipal = Get-AzADServicePrincipal -ObjectId $servicePrincipal.Id
        if (-not $servicePrincipal) {
            throw "SERVICEPRINCIPAL_GET_ERROR"
        }
  
        #Create AD Application Secret
        $startDate = $(Get-Date)
        $endDate = $startDate.AddYears(2)
        Write-Host "Connect AzureAD..." -ForegroundColor White -BackgroundColor Black
        Connect-AzureAD -TenantId "$($TenantId)" | Out-Null
        $secret = New-AzureADApplicationPasswordCredential -ObjectId $aadApplication.ObjectId `
            -CustomKeyIdentifier "DevOps" `
            -StartDate $startDate -EndDate $endDate

        return @(
            $aadApplication, $secret.Value, $servicePrincipal.ApplicationId
        )
    }
    return @(
        "***", "***", "***"
    )
}

function New-AzDevOpsProject {
    param (
        [string]$Version = "5.1",
        [Parameter(Mandatory = $true)]
        [string]$Organization = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$PAT = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$Body = [string]::Empty
    )
    
    $encodedPat = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$PAT"))

    $data = $Body | ConvertFrom-Json
    $resp = Invoke-WebRequest -Uri "https://dev.azure.com/$($Organization)/_apis/projects?api-version=$($Version)" `
        -Method "POST" `
        -Headers @{Authorization = "Basic $encodedPat" } `
        -ContentType application/json `
        -Body $Body

    $returnCode = $resp.StatusCode
    $returnStatus = $resp.StatusDescription

    $response = $resp | ConvertFrom-Json
    if ($returnCode -ne "202") {
        Write-Host "Create project failed - $returnCode $returnStatus" -ForegroundColor White -BackgroundColor Red
        break
    }
    else {
        Write-Host "Project $($Data.Name)($($response.id)) is created."
    }
    return $response
    
}

function Get-AzDevOpsProject {
    param (
        [string]$Version = "5.1",
        [Parameter(Mandatory = $true)]
        [string]$Organization = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$PAT = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$ProjectName = [string]::Empty
    )
    $encodedPat = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$PAT"))

    $resp = try {
        (Invoke-WebRequest -Uri "https://dev.azure.com/$($Organization)/_apis/projects/$($ProjectName)?api-version=$($Version)" `
                -Method "GET" `
                -Headers @{Authorization = "Basic $encodedPat" } `
                -ContentType application/json ).BaseResponse

    }
    catch {
        Write-Verbose "An exception was caught: $($_.Exception.Message)"
        $_.Exception.Response 
    }

    return $resp | ConvertTo-Json
    
}

function Get-AzDevOpsServiceConnection {
    param (
        [string]$Version = "5.1-preview.2",
        [Parameter(Mandatory = $true)]
        [string]$Organization = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$PAT = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$ProjectName = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$EndPointName = [string]::Empty
    )
    $encodedPat = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$PAT"))

    $resp = try {
        (Invoke-WebRequest -Uri "https://dev.azure.com/$($Organization)/$($ProjectName)/_apis/serviceendpoint/endpoints?endpointNames=$($EndPointName)&api-version=$($Version)" `
                -Method "GET" `
                -Headers @{Authorization = "Basic $encodedPat" } `
                -ContentType application/json )

    }
    catch {
        Write-Verbose "An exception was caught: $($_.Exception.Message)"
        $_.Exception.Response 
    }

    return $resp.Content | ConvertTo-Json | ConvertFrom-Json
}

function New-AzDevOpsServiceConnection {
    param (
        [string]$Version = "5.1-preview.2",
        [Parameter(Mandatory = $true)]
        [string]$Organization = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$ProjectName = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$PAT = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$Body = [string]::Empty
    )

    $encodedPat = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$PAT"))

    $data = $Body | ConvertFrom-Json
    $resp = Invoke-WebRequest -Uri "https://dev.azure.com/$($Organization)/$($ProjectName)/_apis/serviceendpoint/endpoints?api-version=$($Version)" `
        -Method "POST" `
        -Headers @{Authorization = "Basic $encodedPat" } `
        -ContentType application/json `
        -Body $Body

    return $resp | ConvertFrom-Json
    
}

function Start-Execution {
    param (
        [string]$organizationName = [string]::Empty,
        [string]$pat = [string]::Empty,
        [string]$projectName = [string]::Empty,
        [string]$projectDescription = [string]::Empty,
        [string]$sourceControlType = "Git",
        [string]$subscriptionId = [string]::Empty,
        [string]$tenantId = [string]::Empty
    )

    $createProjectBody = @"
{
    "name": "$($projectName)",
    "description": "$($projectDescription)",
    "capabilities": {
    "versioncontrol": {
        "sourceControlType": "$($sourceControlType)"
    },
    "processTemplate": {
        "templateTypeId": "6b724908-ef14-45cf-84f8-768b5384da45"
    }
    }
}
"@
    Write-Host "Checking modules..."
    Initialize-Modules
    Write-Host "Checking AD Application..."
    $appResult = New-AzADAppForDevops -OrganizationName "$($organizationName)" `
        -ApplicationName "$($projectName)-DevOps" `
        -TenantId "$($tenantId)" `
        -SubscriptionId "$($subscriptionId)"
    
    $servicePrincipalId = $appResult[3]
    $servicePrincipalKey = $appResult[2]
    Write-Host "Checking if project $($projectName) exists ..." -ForegroundColor White -BackgroundColor Black
    $project = Get-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" -ProjectName "$($projectName)"
    $result = $project | ConvertFrom-Json
    
    if ($result.StatusCode -eq 404) {
        Write-Host "Project($($projectName)) is not existing, creating a project ..." -ForegroundColor White -BackgroundColor Black
        New-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" -Body $createProjectBody
        Write-Host "Project($($projectName)) is created ..." -ForegroundColor White -BackgroundColor Black
        $project = Get-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" -ProjectName "$($projectName)"    
    }elseif($result.StatusCode -eq 401){
        Write-Host "Unauthorized" -ForegroundColor Red -BackgroundColor Black
        exit -1;
    }
    Write-Host "Project($($projectName)) is ready..." -ForegroundColor White -BackgroundColor Black
    $addServicePrincipalBody = @"
{
    "data": {
        "SubscriptionId": "$($subscriptionId)",
        "SubscriptionName": "Visual Studio Premium with MSDN"
    },
    "name": "$($projectName)-ServiceEndpoint",
    "type": "AzureRM",
    "authorization": {
        "parameters": {
            "tenantId": "$($tenantId)",
            "servicePrincipalId": "$($servicePrincipalId)",
            "servicePrincipalKey": "$($servicePrincipalKey)"
        },
        "scheme": "ServicePrincipal"
    },
    "isReady": true
}
"@
    Write-Host "Checking Service Connection ..." -ForegroundColor White -BackgroundColor Black
    $serviceConnection = Get-AzDevOpsServiceConnection -Organization "$($organizationName)" `
        -PAT "$($pat)" `
        -ProjectName "$($projectName)" `
        -EndPointName "$($projectName)-ServiceEndpoint"

    if (0 -eq ($serviceConnection | ConvertFrom-Json).count) {
        Write-Host "No service connection, creating a new one ..." -ForegroundColor White -BackgroundColor Black
        New-AzDevOpsServiceConnection -Organization "$($organizationName)" `
            -PAT "$($pat)" `
            -ProjectName "$($projectName)" `
            -Body $addServicePrincipalBody
        Write-Host "Service Connection is created..." -ForegroundColor White -BackgroundColor Black
    }
    Write-Host "Service Connection is ready ..." -ForegroundColor White -BackgroundColor Black
}

Start-Execution -organizationName $OrganizationName `
    -pat $PAT `
    -projectName $ProjectName `
    -projectDescription $ProjectDescription `
    -subscriptionId $SubscriptionId `
    -tenantId $TenantId 

