[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $Path = "C:\temp"
)

Push-Location $Path

# turns out this is super easy. barely an inconvenience!
# make sure NuGet is installed and updated
$progressPreference = 'silentlyContinue'
Write-Verbose "Update-WingetApps - Installing WinGet PowerShell module from PSGallery..."
Write-Verbose "Update-WingetApps - Install NuGet"
$currNuget = Get-PackageProvider NuGet
$latestNuget = Find-PackageProvider NuGet
if (-NOT $currNuget -or ($currNuget.Version -ne $latestNuget.Version)) {
    $null = Install-PackageProvider -Name NuGet -Force -Confirm:$false
}

# install/update Microsoft.WinGet.Client
$modName = 'Microsoft.WinGet.Client'
$currWGMod = Get-Module -ListAvailable $modName
$latestWGMod = Find-Module -Name $modName
if (-NOT $currWGMod -or ($currWGMod.Version -ne $latestWGMod.Version)) {
    $null = Install-Module -Name $modName -Force -Repository PSGallery 
}

# now use winget powershell to do all the work!
Write-Verbose "Update-WingetApps - Using Repair-WinGetPackageManager cmdlet to bootstrap WinGet..."
Repair-WinGetPackageManager -AllUsers -Latest 

# update the path, just in case
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

# update applications using winget
Write-Verbose "Update-WingetApps - Running winget. Saving log to: $Path\winget.log"
winget upgrade --all --force --accept-package-agreements --accept-source-agreements *> .\winget.log 


Pop-Location
Write-Verbose "Update-WingetApps - Done."

# for legacy purposes
<### VARIABLES ###
#region

# make sure the download location is there
$null = mkdir "$updatePath" -Force
Set-Location "$updatePath"

# winget repo details
$wingetOwner = 'microsoft'
$wingetRepo = 'winget-cli'

# XAML download details
$xamlURI = 'https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.8.6'
$xamlFile = "microsoft.ui.xaml.$(Get-Date -Format "yyyyMMdd_HHmmss")"

# VClibs download URL
$VCLibsURI = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'

#endregion VARIABLES


### FUNCTIONS ###
#region

# load the common functions
. "$PSScriptRoot\lib\libFunc.ps1"

# checks for WS2022 and updates permissions
function Update-PermissionsWS2022 {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $WingetPath
    )
    # WS2022 build number
    $22Build = 20348

    # get the OS build
    $osBuild = [System.Environment]::OSVersion.Version.Build

    # update permissions when WS2022
    if ($osBuild -eq $22Build) {
        TAKEOWN /F "$WingetPath" /R /A /D Y *> $null
        ICACLS "$WingetPath" /grant Administrators:F /T *> $null
    }
}

# searches the local system for winget
function Find-WinGet {
    $wingetFnd = Get-Command winget.exe -EA SilentlyContinue

    if ( -NOT [string]::IsNullOrEmpty($wingetFnd.Source) ) { 
        # found it! return the path
        Write-Verbose "Winget found: $($wingetFnd.Source)"
        return ($wingetFnd.Source)
    }

    # older server installs put winget elsewhere, so look in the alternate path
    $ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe"
    if ($ResolveWingetPath){
        $WingetPath = $ResolveWingetPath[-1].Path
    }

    if ( -NOT [string]::IsNullOrEmpty($WingetPath) ) { 
        # found it! return the path
        Write-Verbose "Winget found: $WingetPath"
        return $WingetPath
    }

    # return $null when this point is reached, signifying that winget was not found
    Write-Verbose "Winget not found!"
    return $null
}


# installs or updates winget
function Install-Winget {
    # install winget
    $wingetOwner = 'microsoft'
    $wingetRepo = 'winget-cli'
    $wingetFile = 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'

    # TO-DO: Auto-detect the license file name!

    $wingetLicenseFile = '76fba573f02545629706ab99170237bc_License1.xml'
    Write-Verbose "Downloading winget"
    $wingetInstaller = Get-LatestGitHubRelease -Owner $wingetOwner -Repo $wingetRepo -File $wingetFile -Path $updatePath -LicenseFile $wingetLicenseFile

    $WingetPath = Find-WinGet

    if ($WingetPath) {
        Write-Verbose "Install winget from: $wingetInstaller"
        try {
            $null = Add-AppxPackage "$wingetInstaller" -EA Stop
        } catch {
            throw "Failed to install winget: $_"
        }
    } else {
        Write-Verbose "Install winget w/ license from: $wingetInstaller"
        try {
            Write-Verbose "Add-AppxProvisionedPackage -Online -PackagePath `"$wingetInstaller`" -LicensePath `"$updatePath\$wingetLicenseFile`""
            $null = Add-AppxProvisionedPackage -Online -PackagePath "$wingetInstaller" -LicensePath "$updatePath\$wingetLicenseFile" -EA Stop
        } catch {
            throw "Failed to install winget: $_"
        }
    }    

    #Remove-Item .\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle -Force

    # get the updated winget path
    $WingetPath = Find-WinGet

    # update permissions on WS2022
    $wingetDir = Split-Path $WingetPath -Parent
    if ($wingetDir) {
        Update-PermissionsWS2022 $wingetDir
    }

}

#endregion FUNCTIONS


### MAIN ###

# install winget if it's not installed
$wingetFnd = Find-WinGet

if (-NOT $wingetFnd) {
    ## install required components
    # Install VCLibs
    Write-Verbose "Installing VCLibs from: $VCLibsURI"

    $oldProgress = $global:ProgressPreference
    $global:ProgressPreference = 'SilentlyContinue'

    try {
        Write-Verbose "Installing VCLibs from: $VCLibsURI"
        $null = Add-AppxPackage "$VCLibsURI" -EA Stop
    } catch {
        throw "Failed to install a required component, VCLibs. Error: $_"
    }

    # Install Microsoft.UI.Xaml from NuGet
    $xamlPkg = Get-WebFile -URI $xamlURI -Path $updatePath -FileName "$xamlFile.zip"

    if (-NOT (Test-Path "$xamlPkg") ) {
        throw "A required component, XAML, failed to download."
    }

    Expand-Archive "$xamlPkg"
    $xamlAppxPkg = Get-ChildItem ".\$xamlFile\tools\AppX\x64\Release\" -Filter "Microsoft.UI.Xaml*.appx" | ForEach-Object FullName
    try {
        Write-Verbose "Installing XAML from: $xamlAppxPkg"
        $null = Add-AppxPackage "$xamlAppxPkg" -EA Stop
    } catch {
        throw "Failed to install a required component, XAML. Error: $_"
    }
    
    ## install winget
    Install-Winget

    $global:ProgressPreference = $oldProgress

} else {
    # make sure winget is up-to-date
    # get the latest release version
    [version]$wingetLatest = Find-GitHubLatestVersion -Owner $wingetOwner -Repo $wingetRepo 

    # get the current winget version
    $WingetPath = Find-WinGet

    if ( $WingetPath -notin $ENV:PATH.Split(';') ) { $ENV:PATH += ";$WingetPath" }
    [version]$wingetCurrent = winget --version | Select-String -Pattern "(?<ver>\d{1,4}\.\d{1,4}\.\d{1,6})" | ForEach-Object {$_.Matches.Groups[1].Value}

    Write-Verbose "Local version: $($wingetCurrent); Online version: $($wingetLatest)"

    # update if there is a newer version
    if ($wingetLatest -gt $wingetCurrent) {
        # Install the latest release of Microsoft.DesktopInstaller from GitHub
        Install-Winget
    }
}


# make sure winget is in the path
$WingetPath = Find-WinGet
if ( $WingetPath -notin $ENV:PATH.Split(';') ) { $ENV:PATH += ";$WingetPath" }
    
# update permissions on WS2022
$wingetDir = Split-Path $WingetPath -Parent
if ($wingetDir) {
    Update-PermissionsWS2022 $wingetDir
}

# update applications using winget
winget upgrade --all --force --accept-package-agreements --accept-source-agreements *> .\winget.log 
#>
