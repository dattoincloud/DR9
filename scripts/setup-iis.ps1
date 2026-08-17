[CmdletBinding()]
param(
    [string]$SiteName = "DR9ClassicAspDemo",
    [ValidateRange(1024, 65535)]
    [int]$Port = 8088
)

$ErrorActionPreference = "Stop"

if ($PSVersionTable.PSEdition -ne "Desktop" -or $PSVersionTable.PSVersion.Major -lt 5) {
    throw "Use 64-bit Windows PowerShell 5.1, not PowerShell 7."
}
if (-not [Environment]::Is64BitProcess) {
    throw "Open the 64-bit Windows PowerShell host, then run this script again."
}

$currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator

if (-not $currentPrincipal.IsInRole($adminRole)) {
    throw "Run this script from Windows PowerShell as Administrator."
}

$resolvedSitePath = (Resolve-Path -LiteralPath (Split-Path -Parent $PSScriptRoot)).Path
$databasePath = Join-Path $resolvedSitePath "App_Data\tasks.mdb"
$dataPath = Join-Path $resolvedSitePath "App_Data"
$applicationPoolName = $SiteName + "Pool"

$features = @(
    "IIS-WebServerRole",
    "IIS-WebServer",
    "IIS-CommonHttpFeatures",
    "IIS-DefaultDocument",
    "IIS-StaticContent",
    "IIS-HttpErrors",
    "IIS-ApplicationDevelopment",
    "IIS-ASP",
    "IIS-ISAPIExtensions",
    "IIS-RequestFiltering",
    "IIS-AnonymousAuthentication",
    "IIS-HttpLogging",
    "IIS-ManagementScriptingTools",
    "IIS-ManagementConsole"
)

$availableFeatures = @()
try {
    $availableFeatures = @(Get-WindowsOptionalFeature -Online -ErrorAction Stop | Select-Object -ExpandProperty FeatureName)
} catch {
    $availableFeatures = @()
}

$featuresToEnable = @()
foreach ($feature in $features) {
    if ($feature -in $availableFeatures) {
        $featuresToEnable += $feature
    } else {
        Write-Verbose "Skipping unsupported Windows feature: $feature"
    }
}

if ($featuresToEnable.Count -gt 0) {
    $featureResult = Enable-WindowsOptionalFeature -Online -FeatureName $featuresToEnable -All -NoRestart
    if ($featureResult -and $featureResult.RestartNeeded) {
        throw "Windows enabled IIS but requires a restart. Restart, then run this script again."
    }
} else {
    Write-Host "No IIS optional features were available on this Windows image; continuing with the IIS configuration step."
}

Import-Module WebAdministration

function Get-BindingSiteName {
    param($Binding)

    if ($Binding.ItemXPath -match "site\[@name='([^']+)'") {
        return $Matches[1]
    }

    return ""
}

$targetBinding = "127.0.0.1:" + $Port + ":"
$wildcardBinding = "*:" + $Port + ":"
$conflictingBinding = Get-WebBinding -Protocol "http" | Where-Object {
    $ownerName = Get-BindingSiteName -Binding $_
    ($_.bindingInformation -eq $targetBinding -or $_.bindingInformation -eq $wildcardBinding) -and
        $ownerName -ne $SiteName
}
if ($conflictingBinding) {
    throw "Port $Port is already used by another IIS site. Choose a different -Port."
}

if (-not (Test-Path -LiteralPath ("IIS:\AppPools\" + $applicationPoolName))) {
    New-WebAppPool -Name $applicationPoolName | Out-Null
}

Set-ItemProperty ("IIS:\AppPools\" + $applicationPoolName) -Name managedRuntimeVersion -Value ""
Set-ItemProperty ("IIS:\AppPools\" + $applicationPoolName) -Name managedPipelineMode -Value "Integrated"
Set-ItemProperty ("IIS:\AppPools\" + $applicationPoolName) -Name enable32BitAppOnWin64 -Value $true
Set-ItemProperty ("IIS:\AppPools\" + $applicationPoolName) -Name processModel.identityType -Value "ApplicationPoolIdentity"
Set-ItemProperty ("IIS:\AppPools\" + $applicationPoolName) -Name processModel.loadUserProfile -Value $true
Set-ItemProperty ("IIS:\AppPools\" + $applicationPoolName) -Name processModel.maxProcesses -Value 1

$existingSite = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
if (-not $existingSite) {
    New-Website -Name $SiteName -Port $Port -IPAddress "127.0.0.1" -PhysicalPath $resolvedSitePath -ApplicationPool $applicationPoolName | Out-Null
} else {
    $existingPath = [Environment]::ExpandEnvironmentVariables($existingSite.physicalPath)
    if ([IO.Path]::GetFullPath($existingPath) -ne [IO.Path]::GetFullPath($resolvedSitePath)) {
        throw "IIS site '$SiteName' already exists with a different physical path: $existingPath"
    }

    Set-ItemProperty ("IIS:\Sites\" + $SiteName) -Name applicationPool -Value $applicationPoolName

    $hasTargetBinding = Get-WebBinding -Name $SiteName -Protocol "http" | Where-Object {
        $_.bindingInformation -eq $targetBinding -or $_.bindingInformation -eq $wildcardBinding
    }
    if (-not $hasTargetBinding) {
        New-WebBinding -Name $SiteName -Protocol "http" -IPAddress "127.0.0.1" -Port $Port | Out-Null
    }
}

Set-WebConfigurationProperty `
    -PSPath "MACHINE/WEBROOT/APPHOST" `
    -Location $SiteName `
    -Filter "system.webServer/security/authentication/anonymousAuthentication" `
    -Name "enabled" `
    -Value $true

Set-WebConfigurationProperty `
    -PSPath "MACHINE/WEBROOT/APPHOST" `
    -Location $SiteName `
    -Filter "system.webServer/security/authentication/anonymousAuthentication" `
    -Name "userName" `
    -Value ""

$appPoolAccount = "IIS AppPool\" + $applicationPoolName
if (-not (Test-Path -LiteralPath $dataPath)) {
    New-Item -ItemType Directory -Path $dataPath | Out-Null
}

$rootAcl = Get-Acl -LiteralPath $resolvedSitePath
$rootRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $appPoolAccount,
    "ReadAndExecute",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)
$rootAcl.SetAccessRule($rootRule)
Set-Acl -LiteralPath $resolvedSitePath -AclObject $rootAcl

$dataAcl = Get-Acl -LiteralPath $dataPath
$dataRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $appPoolAccount,
    "Modify",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)
$dataAcl.SetAccessRule($dataRule)
Set-Acl -LiteralPath $dataPath -AclObject $dataAcl

$cscriptPath = Join-Path $env:WINDIR "SysWOW64\cscript.exe"

if (-not (Test-Path -LiteralPath $databasePath)) {
    & $cscriptPath //nologo (Join-Path $PSScriptRoot "create_database.vbs")
    if ($LASTEXITCODE -ne 0) {
        throw "Database creation failed with exit code $LASTEXITCODE."
    }
}

& $cscriptPath //nologo (Join-Path $PSScriptRoot "smoke_test.vbs")
if ($LASTEXITCODE -ne 0) {
    throw "The selected worker bitness cannot pass the Access database smoke test."
}

if ((Get-WebAppPoolState -Name $applicationPoolName).Value -ne "Started") {
    Start-WebAppPool -Name $applicationPoolName
}
if ((Get-WebsiteState -Name $SiteName).Value -ne "Started") {
    Start-Website -Name $SiteName
}

$siteUrl = "http://127.0.0.1:" + $Port + "/"
try {
    $healthResponse = Invoke-WebRequest -Uri $siteUrl -UseBasicParsing -TimeoutSec 15
    if ($healthResponse.StatusCode -ne 200) {
        throw "Unexpected HTTP status $($healthResponse.StatusCode)."
    }
    if ($healthResponse.Content -notmatch "<title>Legacy Task Board</title>" -or $healthResponse.Content -match "<%@") {
        throw "The response does not look like executed Classic ASP output."
    }
} catch {
    throw "IIS was configured, but the HTTP smoke test failed at $siteUrl $($_.Exception.Message)"
}

$blockedPaths = @(
    "App_Data/tasks.mdb",
    "includes/db.asp",
    "scripts/setup-iis.ps1"
)
foreach ($blockedPath in $blockedPaths) {
    $blockedUrl = $siteUrl + $blockedPath
    try {
        $unexpectedResponse = Invoke-WebRequest -Uri $blockedUrl -UseBasicParsing -TimeoutSec 15
        throw "Sensitive path was served with HTTP $($unexpectedResponse.StatusCode): $blockedUrl"
    } catch {
        if ($null -eq $_.Exception.Response) {
            throw "Could not verify that the sensitive path is blocked: $blockedUrl $($_.Exception.Message)"
        }

        $blockedStatus = [int]$_.Exception.Response.StatusCode
        if ($blockedStatus -ne 403 -and $blockedStatus -ne 404) {
            throw "Sensitive path returned unexpected HTTP $blockedStatus instead of 403/404: $blockedUrl"
        }
    }
}

Write-Host "Classic ASP demo is configured."
Write-Host ("URL: " + $siteUrl)
Write-Host ("Pool: " + $applicationPoolName + " (32-bit, Jet 4.0)")
Write-Host ("Database ACL: Modify for " + $appPoolAccount)
