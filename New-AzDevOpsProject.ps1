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
    $guid = $(New-Guid).ToString()
    $applicationName = "$($OrganizationName)-$ApplicationName"
    $appIdUri = "https://VisualStudio/SPN$($guid)"
    $appReplyUrls = @($appIdUri)
    
    if ($SubscriptionId) {
        Write-Host "Connecting to AzAccount... Please enter sign-in account credantials." -ForegroundColor White -BackgroundColor Black
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
        Write-Host "Connecting to AzureAD...Please enter sign-in account credantials." -ForegroundColor White -BackgroundColor Black
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
        [string]$ProjectName = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$ProjectDescription = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$SourceControlType = "Git"
    )
    
    $createProjectBody = @"
{
    "name": "$($ProjectName)",
    "description": "$($ProjectDescription)",
    "capabilities": {
        "versioncontrol": {
            "sourceControlType": "$($SourceControlType)"
        },
        "processTemplate": {
            "templateTypeId": "6b724908-ef14-45cf-84f8-768b5384da45"
        }
    }
}
"@
    $encodedPat = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$PAT"))

    $data = $createProjectBody | ConvertFrom-Json
    $resp = Invoke-WebRequest -Uri "https://dev.azure.com/$($Organization)/_apis/projects?api-version=$($Version)" `
        -Method "POST" `
        -Headers @{Authorization = "Basic $encodedPat" } `
        -ContentType application/json `
        -Body $createProjectBody

    $returnCode = $resp.StatusCode
    $returnStatus = $resp.StatusDescription

    $response = $resp | ConvertFrom-Json
    if ($returnCode -ne "202") {
        Write-Host "Create project failed - $returnCode $returnStatus" -ForegroundColor White -BackgroundColor Red
        break
    }
    else {
        Write-Host "Project $($data.Name)($($response.id)) is created."
    }
    return $response
    
}

function Get-AzDevOpsProject {
    param (
        [string]$Version = "6.0-preview.4",
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
                -ContentType application/json )

    }
    catch {
        Write-Verbose "An exception was caught: $($_.Exception.Message)"
        $_.Exception.Response 
    }

    if($null -eq $resp.Content)
    {
        return $null
    }
    return $resp.Content | ConvertFrom-Json
    
}

function Get-AzDevOpsServiceConnection {
    param (
        [string]$Version = "6.0-preview.4",
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
        [string]$TenantId = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$SubscriptionId = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$ProjectName = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$ProjectId = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$PAT = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$ServicePrincipalId = [string]::Empty,
        [Parameter(Mandatory = $true)]
        [string]$ServicePrincipalKey = [string]::Empty
    )

    $addServicePrincipalBody = @"
{
    "data": {
        "SubscriptionId": "$($SubscriptionId)",
        "SubscriptionName": "Visual Studio Premium with MSDN",
        "environment": "AzureCloud",
        "scopeLevel": "Subscription",
        "creationMode": "Manual"
    },
    "name": "$($ProjectName)-ServiceEndpoint",
    "type": "AzureRM",
    "authorization": {
        "parameters": {
            "tenantid": "$($TenantId)",
            "serviceprincipalid": "$($ServicePrincipalId)",
            "serviceprincipalkey": "$($ServicePrincipalKey)",
            "authenticationType": "spnKey",
        },
        "scheme": "ServicePrincipal"
    },
    "isReady": true,
    "isShared": false,
    "serviceEndpointProjectReferences": [
        {
          "projectReference": {
            "id": "$($ProjectId)",
            "name": "$($ProjectName)"
          },
          "name": "$($ProjectName)-ServiceEndpoint",
        }
      ]
}
"@

    $encodedPat = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$PAT"))

    $resp = Invoke-WebRequest -Uri "https://dev.azure.com/$($Organization)/$($ProjectName)/_apis/serviceendpoint/endpoints?api-version=$($Version)" `
        -Method "POST" `
        -Headers @{Authorization = "Basic $encodedPat" } `
        -ContentType application/json `
        -Body $addServicePrincipalBody

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



    Write-Host "Checking modules..."
    Initialize-Modules
    Write-Host "Checking AD Application..."
    $appResult = New-AzADAppForDevops -OrganizationName "$($organizationName)" `
        -ApplicationName "$($projectName)-$($subscriptionId)" `
        -TenantId "$($tenantId)" `
        -SubscriptionId "$($subscriptionId)"
    
    $servicePrincipalId = $appResult[3]
    $servicePrincipalKey = $appResult[2]
    Write-Host "Checking if project $($projectName) exists..." -ForegroundColor White -BackgroundColor Black
    $project = Get-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" -ProjectName "$($projectName)"
    $projectid=[string]::Empty
    if ($null -eq $project) {
        Write-Host "Project($($projectName)) is not existing, creating a project..." -ForegroundColor White -BackgroundColor Black
        New-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" `
            -ProjectName "$($projectName)" `
            -ProjectDescription "$($projectDescription)" `
            -SourceControlType "Git"
    }
    $project = Get-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" -ProjectName "$($projectName)" 
    $projectid=$project.id  

    Write-Host "Checking Service Connection..." -ForegroundColor White -BackgroundColor Black
    $serviceConnection = Get-AzDevOpsServiceConnection -Organization "$($organizationName)" `
        -PAT "$($pat)" `
        -ProjectName "$($projectName)" `
        -EndPointName "$($projectName)-ServiceEndpoint"

    if (0 -eq ($serviceConnection | ConvertFrom-Json).count) {
        Write-Host "No service connection, creating a new one..." -ForegroundColor White -BackgroundColor Black
        $serviceConnection = New-AzDevOpsServiceConnection -Organization "$($organizationName)" `
                                -PAT "$($pat)" `
                                -TenantId "$($tenantId)" `
                                -SubscriptionId "$($subscriptionId)" `
                                -ProjectId "$($projectid)" `
                                -ProjectName "$($projectName)" `
                                -ServicePrincipalId "$($servicePrincipalId)" `
                                -ServicePrincipalKey "$($servicePrincipalKey)" `

        Write-Host "Service Connection is created..." -ForegroundColor White -BackgroundColor Black
    }
}

Start-Execution -organizationName $OrganizationName `
    -pat $PAT `
    -projectName $ProjectName `
    -projectDescription $ProjectDescription `
    -subscriptionId $SubscriptionId `
    -tenantId $TenantId 

