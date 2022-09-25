# setup WS2022
# Based on Initialize-NetworkLab

<#

TO-DO:


#>

[CmdletBinding()]
param ()

### CONSTANTS ###
#region

# where to put downloads
$savePath = "C:\Temp"

# list of exact winget app IDs to install
[array]$wingetApps = "JanDeDobbeleer.OhMyPosh", "WiresharkFoundation.Wireshark", "Microsoft.VisualStudioCode", "Git.Git" # "Microsoft.PowerShell", "Microsoft.WindowsTerminal" -- winget tries to use MSIX which fails on 2022, so grab via github

# powershell repro and file extension
$pwshRepo = "PowerShell/PowerShell"
$pwshExt = "msi"

# Windows Terminal details
$termRepo = "Microsoft/Terminal"

# winget repro and file extension
$wingetRepo = "microsoft/winget-cli"
$wingetExt = "msixbundle"

# repro for Caskaydia Cove Nerd Font
$repoCCNF = "ryanoasis/nerd-fonts"

# name of the preferred pretty font, CaskaydiaCove NF
$fontName = "CaskaydiaCove NF"

# the zip file where CC NF is in
$fontFile = "CascadiaCode.zip"

# list of commands to add to the PowerShell profile
[string[]]$profileLines = 'Import-Module -Name Terminal-Icons',
                          'oh-my-posh --init --shell pwsh --config ~/slimfat.omp.json | Invoke-Expression',
                          'cls'

# npcap URL
# Wireshark current version is 1.60, 1.71 is the newest
#$npcapURL = "https://nmap.org/npcap/dist/npcap-1.60.exe"
$npcapURL = 'https://npcap.com/dist/npcap-1.71.exe'


# VCLib URL - needed for Terminal
$vclibUrl = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
$vclibAppxName = 'Microsoft.VCLibs.140.00.UWPDesktop'

# XAML URL for Terminal
# 2022-09-23 - Version 2.7 currently required for winget. NuGet's latest package is XAML 2.8 which causes winget install to fail.
#$xamlUrl = 'https://www.nuget.org/packages/Microsoft.UI.Xaml/'
$xamlUrl = 'https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.7.3'
$xamlAppxName = 'Microsoft.UI.Xaml.2.7'


#endregion CONSTANTS



### FUNCTIONS ###
#region

# FUNCTION: Find-GitRelease
# PURPOSE:  Calls Github API to retrieve details about the latest release. Returns a PSCustomObject with repro, version (tag_name), and download URL.
function Find-GitRelease
{
    [CmdletBinding()]
    param(
        [string]$Repo,
        [switch]$Latest
    )

    Write-Verbose "Find-GitRelease - Begin"

    # make sure we don't try to use an insecure SSL/TLS protocol when downloading files
    $secureProtocols = @() 
    $insecureProtocols = @( [System.Net.SecurityProtocolType]::SystemDefault, 
                            [System.Net.SecurityProtocolType]::Ssl3, 
                            [System.Net.SecurityProtocolType]::Tls, 
                            [System.Net.SecurityProtocolType]::Tls11) 
    foreach ($protocol in [System.Enum]::GetValues([System.Net.SecurityProtocolType])) 
    { 
        if ($insecureProtocols -notcontains $protocol) 
        { 
            $secureProtocols += $protocol 
        } 
    } 
    [System.Net.ServicePointManager]::SecurityProtocol = $secureProtocols

    if ($Latest.IsPresent)
    {
        Write-Verbose "Find-GitRelease - Finding latest release."
        $baseApiUri = "https://api.github.com/repos/$($repo)/releases/latest"
    }
    else
    {
        Write-Verbose "Find-GitRelease - Finding all releases."
        $baseApiUri = "https://api.github.com/repos/$($repo)/releases"
    }

    # get the available releases
    Write-Verbose "Find-GitRelease - Processing repro: $repo"
    Write-Verbose "Find-GitRelease - Making Github API call to: $baseApiUri"
    try 
    {
        if ($pshost.Version.Major -le 5)
        {
            $rawReleases = Invoke-WebRequest $baseApiUri -UseBasicParsing -EA Stop
        }
        elseif ($pshost.Version.Major -ge 6)
        {
            $rawReleases = Invoke-WebRequest $baseApiUri -EA Stop
        }
        else 
        {
            return (Write-Error "Unsupported version of PowerShell...?" -EA Stop)
        }
    }
    catch 
    {
        return (Write-Error "Could not get GitHub releases. Error: $_" -EA Stop)        
    }

    
    if ($Latest.IsPresent)
    {
        try
        {
            [version]$version = ($rawReleases.Content | ConvertFrom-Json).tag_name
        }
        catch
        {
            $version = ($rawReleases.Content | ConvertFrom-Json).tag_name
        }

        $dlURI = ($rawReleases.Content | ConvertFrom-Json).Assets.browser_download_url

        Write-Verbose "Find-GitRelease - Processing latest version."
        $releases = [PSCustomObject]@{
            Repo    = $repo
            Version = $version
            URL     = $dlURI
        }

        Write-Verbose "Find-GitRelease - Found: $version at $dlURI"
    }
    else
    {
        $releases = @()
        Write-Verbose "Find-GitRelease - Processing $($version.Count) versions."
        $jsonReleases = $rawReleases.Content | ConvertFrom-Json
        for ($i = 0; $i -lt $jsonReleases.Count; $i++)
        {
            $tmpObj = [PSCustomObject]@{
                Name    = $jsonReleases[$i].name
                Version = [version]($jsonReleases[$i].tag_name.TrimStart("v"))
                URL     = $jsonReleases[$i].Assets.browser_download_url
            }

            $releases += $tmpObj

            Remove-Variable tmpObj -EA SilentlyContinue
        }
    }

    Write-Verbose "Find-GitRelease - End"
    return $releases
} #end Find-GitRelease



# FUNCTION: Get-WebFile
# PURPOSE:  
function Get-WebFile
{
    [CmdletBinding()]
    param ( 
        [string]$URI,
        [string]$savePath,
        [string]$fileName
    )

    Write-Verbose "Get-WebFile - Begin"
    Write-Verbose "Get-WebFile - Attempting to download: $dlUrl"

    # make sure we don't try to use an insecure SSL/TLS protocol when downloading files
    $secureProtocols = @() 
    $insecureProtocols = @( [System.Net.SecurityProtocolType]::SystemDefault, 
                            [System.Net.SecurityProtocolType]::Ssl3, 
                            [System.Net.SecurityProtocolType]::Tls, 
                            [System.Net.SecurityProtocolType]::Tls11) 
    foreach ($protocol in [System.Enum]::GetValues([System.Net.SecurityProtocolType])) 
    { 
        if ($insecureProtocols -notcontains $protocol) 
        { 
            $secureProtocols += $protocol 
        } 
    } 
    [System.Net.ServicePointManager]::SecurityProtocol = $secureProtocols

    try 
    {
        Invoke-WebRequest -Uri $URI -OutFile "$savePath\$fileName"
    } 
    catch 
    {
        return (Write-Error "Could not download $URI`: $($Error[0].ToString())" -EA Stop)
    }

    Write-Verbose "Get-WebFile - File saved to: $savePath\$fileName"
    Write-Verbose "Get-WebFile - End"
    return "$savePath\$fileName"
} #end Get-WebFile



function Install-FromGithub
{
    [CmdletBinding()]
    param (
        $repo,
        $extension,
        $savePath
    )

    # appx extensions
    $appxExt = "msixbundle", "appx"

    # download wt
    $release = Find-GitRelease $repo -Latest
    $fileName = "$(($repo -split '/')[-1])`.$extension"

    # find the URL
    $URL = $release.URL | Where-Object { $_ -match "^.*.$extension$" }

    if ($URL -is [array])
    {
        # try to find the URL based on architecture (x64/x86) and OS (Windows)
        if ([System.Environment]::Is64BitOperatingSystem)
        {
            $osArch = "x64"
        }
        else 
        {
            $osArch = "x86"
        }
                
        $URL = $URL | Where-Object { $_ -match $osArch }

        if ($URL -is [array])
        {
            $URL = $URL | Where-Object { $_ -match "win" }

            # can do more, but this is good enough for this script
        }
    }

    try 
    {
        $installFile = Get-WebFile -URI $URL -savePath $savePath -fileName $fileName -EA Stop
        
        if ($extension -in  $appxExt)
        {
            Add-AppxPackage $installFile -EA Stop
        }
        elseif ($extension -eq "msi" -and $fileName -match "powershell")
        {
            msiexec.exe /package "$installFile" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1
        }
        else
        {
            Start-Process "$installFile" -Wait
        }
    }
    catch 
    {
        return (Write-Error "$_" -EA Stop)        
    }
}

function Get-InstalledFonts
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    return ((New-Object System.Drawing.Text.InstalledFontCollection).Families)
}

function Install-Font
{
    [CmdletBinding()]
    param (
        [Parameter()]
        $Path
    )

    $FONTS = 0x14
    $CopyOptions = 4 + 16;
    $objShell = New-Object -ComObject Shell.Application
    $objFolder = $objShell.Namespace($FONTS)

    foreach ($font in $Path)
    {
        $CopyFlag = [String]::Format("{0:x}", $CopyOptions);
        $objFolder.CopyHere($font.fullname,$CopyFlag)
    }
}

#endregion FUNCTIONS




### MAIN ###

$null = mkdir $savePath -EA SilentlyContinue


## install newest pwsh release ##
try
{
    Install-FromGithub -repo $pwshRepo -extension $pwshExt -savePath $savePath
}
catch
{
    return (Write-Error "Failed to download or install PowerShell 7+: $_" -EA Stop)
}

## install pre-req's ##
# install VCLib
try
{
    $vclibFilename = 'Microsoft.VCLibs.x64.14.00.Desktop.appx'
    $vclibFile = Get-WebFile -URI $vclibUrl -savePath $savePath -fileName $vclibFilename -EA Stop
    Add-AppxPackage $vclibFile -EA Stop
}
catch
{
    return (Write-Error "VCLib download or install failed: $_" -EA Stop)
}

# install Microsoft.UI.Xaml
$xamlPage = Invoke-WebRequest $xamlUrl -UseBasicParsing
$xamlDlUrl = $xamlPage.Links | Where-Object { $_.outerHTML -match "outbound-manual-download" } | ForEach-Object { $_.href }

try
{
    $xamlFilename = 'xaml.zip'
    $xamlFile = Get-WebFile -URI $xamlDlUrl -savePath $savePath -fileName $xamlFilename -EA Stop
    
    Expand-Archive $xamlFile -Force -EA Stop

    # find the x64 installer
    $xamlAppx = Get-ChildItem .\xaml -Recurse -Filter "Microsoft.UI.Xaml.*.appx" | Where-Object { $_.FullName -match "x64" }
    
    Add-AppxPackage $xamlAppx -EA Stop
}
catch
{
    return (Write-Error "Microsoft.UI.Xaml download or install failed: $_" -EA Stop)
}

## install Windows Terminal ##
# get the msixbundle for Win10
$termBundles = Find-GitRelease $termRepo | Where-Object { $_.Name -notmatch "Preview"} | Sort-Object -Property Version -Descending
$termURL = $termBundles[0].URL | Where-Object {$_ -match '^*.msixbundle$' -and $_ -match "Win10"}

try 
{
    $termFile = Get-WebFile -URI $termURL -savePath $savePath -fileName "terminal.msixbundle" -EA Stop
    Add-AppxPackage "$termFile" -EA Stop
}
catch 
{
    return (Write-Error "Failed to download or install Windows Terminal: $_" -EA Stop)
}



## download and install winget from github ##
try 
{
    $release = Find-GitRelease $wingetRepo -Latest
    $fileName = "$(($wingetRepo -split '/')[-1])`.$wingetExt"

    # find the URL
    $URL = $release.URL | Where-Object { $_ -match "^.*.$wingetExt$" }

    $installFile = Get-WebFile -URI $URL -savePath $savePath -fileName $fileName -EA Stop

    $URL2 = $release.URL | Where-Object { $_ -match "^.*.xml$" }
    $fileName2 = Split-Path $url2 -Leaf

    $licenseFile = Get-WebFile -URI $URL2 -savePath $savePath -fileName $fileName2 -EA Stop

    Add-AppxProvisionedPackage -Online -PackagePath $installFile -LicensePath $licenseFile -Verbose -EA Stop
}
catch
{
    return (Write-Error "Winget download failed: $_" -EA Stop)
}

# wait for winget to appear in the path
$count = 0
do
{
    Start-Sleep 1
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 
    $wingetFnd = Get-Command winget -EA SilentlyContinue
    
    # 2022-09-23 - Using Add-AppxPackage seems to work with winget on WS2022, so try that is Add-AppxProvisionedPackage didn't see to work
    if ($count -gt 0 -and -NOT $wingetFnd)
    {
        Add-AppxPackage $installFile
    }
    
    $count++
} until ($wingetFnd -or $count -ge 10)

## install winget apps ##
if ($wingetFnd)
{
    foreach ($app in $wingetApps)
    {
        # install things
        winget install $app --exact --accept-package-agreements --accept-source-agreements --silent
    }

    # pwsh doesn't always install the first time. test and retry
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 
    $isPwshFnd = Get-Command pwsh -EA SilentlyContinue

    if (-NOT $isPwshFnd)
    {
        winget install microsoft.powershell

        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 
        $isPwshFnd = Get-Command pwsh -EA SilentlyContinue
        if (-NOT $isPwshFnd)
        {
            return (Write-Error "PowerShell 7+ installation failed. Please install manually and try again." -EA Stop)
        }

    }
}
else
{
    return (Write-Error "Winget installation failed: Winget not found." -EA Stop)
}


## configure PowerShell on Windows Terminal ##
# get CaskaydiaCove NF if not installed
if ($fontName -notin (Get-InstalledFonts))
{
    Write-Verbose "Installing $fontName"
    # get newest font
    $ccnf = Find-GitRelease -repo $repoCCNF -Latest

    # find the correct URL
    $ccnfURL = $ccnf.URL | Where-Object {$_ -match $fontFile}

    # download
    try 
    {
        $ccnfZip = Get-WebFile -URI $ccnfURL -savePath $savePath -fileName $fontFile    
    }
    catch 
    {
        Write-Error "Failed to download $fontFile. Please download and install $fontName manually, or the Nerd Font of your choice."
    }
    
    # extract
    $extractPath = "$savePath\ccnf"
    Expand-Archive -Path $ccnfZip -DestinationPath $extractPath -Force

    # install fonts
    # 2022-09-23 - Added search for OTF files, as TTF support seems to be waning.
    Install-Font (Get-ChildItem "$extractPath" -EA SilentlyContinue | Where-Object Extension -match '\.otf|\.ttf')

    Start-Sleep 30
}

## install modules ##

[scriptblock]$cmd = {
    $nugetVer = Get-PackageProvider -ListAvailable -EA SilentlyContinue | Where-Object Name -match "NuGet" | ForEach-Object { $_.Version }
    [version]$minNugetVer = "2.8.5.208"
    if ($nugetVer -lt $minNugetVer -or $null -eq $nugetVer)
    {
        Write-Verbose "Installing NuGet update."
        $null = Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Force
    }

    # get module(s)
    Install-Module -Name Terminal-Icons,oh-my-posh -Repository PSGallery -Scope CurrentUser -Force
}

# update $env:Path so pwsh will be found
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 

# install pwsh modules
pwsh -NoLogo -NoProfile -Command $cmd



# update the pwsh profile
$pwshProfilePath = pwsh -NoLogo -NoProfile -Command { $PROFILE } 

$pwshProfile = Get-Content $pwshProfilePath -EA SilentlyContinue

if (-NOT (Test-path $pwshProfilePath)) { $null = New-Item $pwshProfilePath -ItemType File -Force }

foreach ($line in $profileLines)
{
    if ($line -notin $pwshProfile)
    {
        $line | Out-File "$pwshProfilePath" -Append -Force
    }
}

# assume WT is installed at this point
# launch WT once to make sure settings.json is generated
Start-Process wt -ArgumentList "-p PowerShell" -WindowStyle Minimized
Start-Sleep 10
Get-Process WindowsTerminal | Stop-Process -Force

$appxPack = Get-AppxPackage -Name "Microsoft.WindowsTerminal" -EA SilentlyContinue
$wtAppData = "$ENV:LOCALAPPDATA\Packages\$($appxPack.PackageFamilyName)\LocalState"

# export settings.json
# clean up comment lines to prevent issues with older JSON parsers (looking at you Windows PowerShell)
try 
{
    $wtJSON =  Get-Content "$wtAppData\settings.json" -EA Stop | Where-Object { $_ -notmatch "^.*//.*$" -and $_ -ne "" -and $_ -ne $null} | ConvertFrom-Json    
}
catch 
{
    return (Write-Error "Failed to update Windows Terminal settings." -EA Stop)
}

# change the font for PowerShell
if ($null -ne $wtJSON.profiles.list.font.face)
{
    $wtJSON.profiles.list | Where-Object { $_.Name -eq "PowerShell" } | ForEach-Object { $_.Font.Face = $fontName }
}
else 
{
    $pwshProfile = $wtJSON.profiles.list | Where-Object { $_.Name -eq "PowerShell" }
    $pwshProfile | Add-Member -NotePropertyName font -NotePropertyValue ([PSCustomObject]@{face="$fontName"})
}

# set PowerShell (pwsh) to the default profil
$pwshGUID = $wtJSON.profiles.list | Where-Object Name -eq "PowerShell" | ForEach-Object { $_.guid }

if ($pwshGUID)
{
    $defaultGUID = $wtJSON.defaultProfile

    if ($defaultGUID -ne $pwshGUID)
    {
        $wtJSON.defaultProfile = $pwshGUID
    }
}

# change some WT defaults... 
$evilPasteSettings = 'largePasteWarning', 'multiLinePasteWarning'

foreach ($imp in $evilPasteSettings)
{
    if ($null -eq $wtJSON."$imp" -or $wtJSON."$imp" -eq $true)
    {
        $wtJSON | Add-Member -NotePropertyName $imp -NotePropertyValue $false -Force
    }
}

# maximize wt on start
$wtJSON | Add-Member -NotePropertyName "launchMode" -NotePropertyValue "maximized" -Force


# save settings
$wtJSON | ConvertTo-Json -Depth 20 | Out-File "$wtAppData\settings.json" -Force -Encoding utf8



# get npcap
try 
{
    $npcapFile = Get-WebFile -URI $npcapURL -savePath $savePath -fileName npcap.exe
    Start-Process "$savePath\npcap.exe" -ArgumentList "/winpcap_mode=disabled" -Wait

    $null = Remove-Item $npcapFile -Force
}
catch 
{
    Write-Warning "Failed to download and install npcap. Please manually download and install: $_"
}

# ALL: Set all network connections to Private
Get-NetConnectionProfile | Where-Object NetworkCategory -eq Public | Set-NetConnectionProfile -NetworkCategory Private

# ALL: Enable File and Printer Sharing on the firewall for Private
Get-NetFirewallRule -DisplayGroup "File and Printer Sharing" | Where-Object { $_.Profile -eq "Private" -or $_.Profile -eq "Any" } | Enable-NetFirewallRule


# update and reboot
Install-PackageProvider -Name NuGet -Force
Install-Module -Name PSWindowsUpdate -MinimumVersion 2.2.0 -Force
Get-WindowsUpdate -AcceptAll -Verbose -WindowsUpdate -Install -AutoReboot