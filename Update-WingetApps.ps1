[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $updatePath = "C:\temp\updates"
)


### VARIABLES ###
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

    # fix potential file permissions for server
    # TO-DO: Detect Server
    if ($WingetPath) {
        TAKEOWN /F "$WingetPath" /R /A /D Y *> $null
        ICACLS "$WingetPath" /grant Administrators:F /T *> $null
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
    [version]$wingetCurrent = winget --version | Select-String -Pattern "^.*(?<ver>\d{1,4\.\d{1,4}\.\d{1,6}).*$" | ForEach-Object {$_.Matches.Groups[1].Value}

    # update if there is a newer version
    if ($wingetLatest -gt $wingetCurrent) {
        # Install the latest release of Microsoft.DesktopInstaller from GitHub
        Install-Winget
    }
}


# make sure winget is in the path
$WingetPath = Find-WinGet
if ( $WingetPath -notin $ENV:PATH.Split(';') ) { $ENV:PATH += ";$WingetPath" }



# update applications using winget
winget upgrade --all --force --accept-package-agreements --accept-source-agreements *> .\winget.log 

