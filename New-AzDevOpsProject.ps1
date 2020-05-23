[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $TenantId = [string]::Empty,
    [Parameter()]
    [string]
    $SubscriptionId = [string]::Empty,
    [Parameter()]
    [string]
    $ServicePrincipalId = [string]::Empty,
    [Parameter()]
    [string]
    $ServicePrincipalKey = [string]::Empty,
    [Parameter()]
    [string]
    $OrganizationName = [string]::Empty,
    [Parameter()]
    [string]
    $PAT = [string]::Empty,
    [Parameter()]
    [string]
    $ProjectName = [string]::Empty,
    [Parameter()]
    [string]
    $ProjectDescription = "Project created with Azure DevOps Rest API"
)

function New-AzDevOpsProject {
    param (
        [string]$Version = "5.1",
        [string]$Organization = [string]::Empty,
        [string]$PAT = [string]::Empty,
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
        [string]$Organization = [string]::Empty,
        [string]$PAT = [string]::Empty,
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
        # @{
        #     IsSuccessStatusCode = $_.Exception.Response.IsSuccessStatusCode
        #     StatusCode          = $_.Exception.Response.StatusCode
        # }
        #$response.BaseResponse.StatusCode.Value__
    }

    # $resp = (Invoke-WebRequest -Uri "https://dev.azure.com/$($Organization)/_apis/projects/$($ProjectName)?api-version=$($Version)" `
    #         -Method "GET" `
    #         -Headers @{Authorization = "Basic $encodedPat" } `
    #         -ContentType application/json ).BaseResponse
    $resp 
    return $resp | ConvertTo-Json
    
}

function Get-AzDevOpsServiceConnection {
    param (
        [string]$Version = "5.1-preview.2",
        [string]$Organization = [string]::Empty,
        [string]$PAT = [string]::Empty,
        [string]$ProjectName = [string]::Empty,
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
        [string]$Organization = [string]::Empty,
        [string]$ProjectName = [string]::Empty,
        [string]$PAT = [string]::Empty,
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
        [string]$tenantId = [string]::Empty,
        [string]$servicePrincipalId = [string]::Empty,
        [string]$servicePrincipalKey = [string]::Empty
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
    Write-Host "Checking if project $($projectName) exists ..." -ForegroundColor White -BackgroundColor Black
    $project = Get-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" -ProjectName "$($projectName)"
    if ($project.IsSuccessStatusCode -eq $false) {
        Write-Host "Project($($projectName)) is not existing, creating a project ..." -ForegroundColor White -BackgroundColor Black
        New-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" -Body $createProjectBody
        Write-Host "Project($($projectName)) is created ..." -ForegroundColor White -BackgroundColor Black
        $project = Get-AzDevOpsProject -Organization "$($organizationName)" -PAT "$($pat)" -ProjectName "$($projectName)"    
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

   
    if (0 -eq  ($serviceConnection | ConvertFrom-Json).count) {
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
    -tenantId $TenantId `
    -servicePrincipalId $ServicePrincipalId `
    -servicePrincipalKey $ServicePrincipalKey `

