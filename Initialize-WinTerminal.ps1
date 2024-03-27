# setup WS2022
# Based on Initialize-NetworkLab
#requires -RunAsAdministrator

<#

TO-DO:
    - Test
    - Add dl of background images
        - with opacity

    - Add DL of custom profile settings

#>

[CmdletBinding()]
param (
    # Set the terminal to demo mode
    [Parameter()]
    [switch]
    $Demo,

    # Installs WSL2
    [Parameter()]
    [switch]
    $WSL
)


Write-Verbose "Begin!"
$script:gDemo = $Demo
if ($gDemo.IsPresent) { Write-Verbose "Demo mode engaged." }
if ($WSL.IsPresent) { Write-Verbose "WSL mode engaged." }

### FUNCTIONS ###
#region

# FUNCTION: Find-GitRelease
# PURPOSE:  Calls Github API to retrieve details about the latest release. Returns a PSCustomObject with repro, version (tag_name), and download URL.
function Find-GitRelease {
    [CmdletBinding()]
    param (
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
    foreach ($protocol in [System.Enum]::GetValues([System.Net.SecurityProtocolType])) { 
        if ($insecureProtocols -notcontains $protocol) { 
            $secureProtocols += $protocol 
        } 
    } 
    [System.Net.ServicePointManager]::SecurityProtocol = $secureProtocols

    if ($Latest.IsPresent) {
        Write-Verbose "Find-GitRelease - Finding latest release."
        $baseApiUri = "https://api.github.com/repos/$($repo)/releases/latest"
    } else {
        Write-Verbose "Find-GitRelease - Finding all releases."
        $baseApiUri = "https://api.github.com/repos/$($repo)/releases"
    }

    # get the available releases
    Write-Verbose "Find-GitRelease - Processing repro: $repo"
    Write-Verbose "Find-GitRelease - Making Github API call to: $baseApiUri"
    try {
        if ($pshost.Version.Major -le 5) {
            $rawReleases = Invoke-WebRequest $baseApiUri -UseBasicParsing -EA Stop
        } elseif ($pshost.Version.Major -ge 6) {
            $rawReleases = Invoke-WebRequest $baseApiUri -EA Stop
        } else {
            return (Write-Error "Unsupported version of PowerShell...?" -EA Stop)
        }
    } catch {
        return (Write-Error "Could not get GitHub releases. Error: $_" -EA Stop)        
    }

    
    if ($Latest.IsPresent) {
        try {
            [version]$version = ($rawReleases.Content | ConvertFrom-Json).tag_name
        } catch {
            $version = ($rawReleases.Content | ConvertFrom-Json).tag_name
        }

        $dlURI = ($rawReleases.Content | ConvertFrom-Json).Assets.browser_download_url

        Write-Verbose "Find-GitRelease - Processing latest version."
        $releases = [PSCustomObject]@{
            Repo    = "$repo"
            Version = "$version"
            URL     = "$dlURI"
        }

        Write-Verbose "Find-GitRelease - Found: $version at $dlURI"
    } else {
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
function Get-WebFile {
    [CmdletBinding()]
    param ( 
        [string]$URI,
        [string]$savePath,
        [string]$fileName
    )

    Write-Verbose "Get-WebFile - Begin"
    Write-Verbose "Get-WebFile - Attempting to download: $URI"

    # make sure we don't try to use an insecure SSL/TLS protocol when downloading files
    $secureProtocols = @() 
    $insecureProtocols = @( [System.Net.SecurityProtocolType]::SystemDefault, 
                            [System.Net.SecurityProtocolType]::Ssl3, 
                            [System.Net.SecurityProtocolType]::Tls, 
                            [System.Net.SecurityProtocolType]::Tls11) 
    foreach ($protocol in [System.Enum]::GetValues([System.Net.SecurityProtocolType])) { 
        if ($insecureProtocols -notcontains $protocol) { 
            $secureProtocols += $protocol 
        } 
    } 
    [System.Net.ServicePointManager]::SecurityProtocol = $secureProtocols

    try {
        Invoke-WebRequest -Uri $URI -OutFile "$savePath\$fileName"
    } catch {
        return (Write-Error "Could not download $URI`: $($Error[0].ToString())" -EA Stop)
    }

    Write-Verbose "Get-WebFile - File saved to: $savePath\$fileName"
    Write-Verbose "Get-WebFile - End"
    return "$savePath\$fileName"
} #end Get-WebFile

function Install-FromGithub {
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
    $URL = $release.URL.Split(" ").Trim(" ") | Where-Object { $_ -match "^.*.$extension$" }

    if ($URL -is [array]) {
        # try to find the URL based on architecture (x64/x86) and OS (Windows)
        if ([System.Environment]::Is64BitOperatingSystem) {
            $osArch = "x64"
        } else {
            $osArch = "x86"
        }
                
        $URL = $URL | Where-Object { $_ -match $osArch }

        if ($URL -is [array]) {
            $URL = $URL | Where-Object { $_ -match "win" }

            # can do more, but this is good enough for this script
        }
    }

    try {
        $installFile = Get-WebFile -URI $URL -savePath $savePath -fileName $fileName -EA Stop
        
        if ($extension -in  $appxExt) {
            Add-AppxPackage $installFile -EA Stop
        } elseif ($extension -eq "msi" -and $fileName -match "powershell") {
            msiexec.exe /package "$installFile" /quiet ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1
        } else {
            Start-Process "$installFile" -Wait
        }
    } catch {
        return (Write-Error "$_" -EA Stop)        
    }
}

function Get-InstalledFonts {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    return ((New-Object System.Drawing.Text.InstalledFontCollection).Families)
}

function Install-Font {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Path
    )

    $FONTS = 0x14
    $CopyOptions = 4 + 16;
    $objShell = New-Object -ComObject Shell.Application
    $objFolder = $objShell.Namespace($FONTS)

    foreach ($font in $Path) {
        $CopyFlag = [String]::Format("{0:x}", $CopyOptions);
        $objFolder.CopyHere($font.fullname,$CopyFlag)
    }
}

#endregion FUNCTIONS

### CONSTANTS ###
#region

# where to put downloads
$savePath = "C:\Temp"

# list of exact winget app IDs to install
[array]$wingetApps = "JanDeDobbeleer.OhMyPosh", "WiresharkFoundation.Wireshark", "Microsoft.VisualStudioCode", "Git.Git"

# powershell repro and file extension
$pwshRepo = "PowerShell/PowerShell"
$pwshExt = "msi"

# Windows Terminal details
$termRepo = "Microsoft/Terminal"
$termPRHint = "PreinstallKit"
$termPRFile = "term-pr.zip"

# VCLibs Desktop is required by winget
$vclibsURL = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
$vclibsFile = "vclibs.zip"

# winget repro and file extension
$wingetRepo = "microsoft/winget-cli"
$wingetExt = "msixbundle"

# repro for Caskaydia Cove Nerd Font
$repoCCNF = "ryanoasis/nerd-fonts"

# name of the preferred pretty font, CaskaydiaCove NF
$fontName = "CaskaydiaCove Nerd Font"

# the zip file where CC NF is in
$fontFile = "CascadiaCode.zip"

# npcap URL
# Wireshark required npcap
$npcapURL = 'https://npcap.com/dist/npcap-1.75.exe'



## DEMO MODE ##
$pwshMods = @()
if ($gDemo.IsPresent) {
    Write-Verbose "Demo presets are being used."
    # font size
    $demoFntSz = 16
    $demoFntWeight = "Normal"

    # set the demo path to the <local desktop>\Demo folder
    $demoPath = "$([System.Environment]::GetFolderPath("Desktop"))\Demo"

    # create the demo path
    $null = mkdir "$demoPath" -Force
    
    # list of commands to add to the PowerShell profile
    [string[]]$profileLines = 'Import-Module -Name Terminal-Icons',
                              'oh-my-posh --init --shell pwsh --config $env:POSH_THEMES_PATH/material.omp.json | Invoke-Expression',
                              "New-PSDrive -Name Demo -PSProvider FileSystem -Root '$demoPath'",
                              "Set-Location Demo:\",
                              'cls'
    
    # PowerShell modules to install
    $pwshMods = "Terminal-Icons", "oh-my-posh", "Az"
} else {
    Write-Verbose "Normal presets are being used."
    # list of commands to add to the PowerShell profile
    [string[]]$profileLines = 'Import-Module -Name Terminal-Icons',
                              'oh-my-posh --init --shell pwsh --config $env:POSH_THEMES_PATH/tokyo.omp.json | Invoke-Expression',
                              'cls'

    # PowerShell modules to install                              
    $pwshMods = "Terminal-Icons", "oh-my-posh"
}


#endregion CONSTANTS


### MAIN ###

<# only works for Windows right now
if ( -NOT $IsWindows ) {
    return (Write-Error "Only Windows is supported at the moment." -EA Stop)
}#>

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (2)" }

# detect Windows version
$rawOS = [System.Environment]::OSVersion
Write-Verbose "Windows Version: $($rawOS.VersionString)"

# bail is OS is less than Win10 19041
if ( $rawOS.Version.Build -lt 19041 ) {
    return (Write-Error "This version of Windows is not supported. The version must be Windows 10 21H2 and Windows Server 2022 or newer." -EA Stop)
}


# create the savePath
Write-Verbose "Create savePath: $savePath"
$null = mkdir $savePath -EA SilentlyContinue


## install newest pwsh release ##
$pwshFnd = Get-Command pwsh -EA SilentlyContinue
[version]$pwshVer = Find-GitRelease -Repo $pwshRepo -Latest | ForEach-Object { $_.Version.TrimStart('v') }

if ( -NOT $pwshFnd -or $pwshVer -gt $pwshFnd.Version ) {
    try
    {
        Write-Verbose "Installing the newest release of PowerShell."
        Install-FromGithub -repo $pwshRepo -extension $pwshExt -savePath $savePath
        Write-Verbose "PowerShell install complete."
    }
    catch
    {
        return (Write-Error "Failed to download or install PowerShell 7+: $_" -EA Stop)
    }
} else {
    Write-Verbose "PowerShell is already installed"
}

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (3)" }

# install Windows Terminal
# this step is only needed for pre-Win11, less than Win10 19041 is covered above.
# install if wt.exe is not found
$wtFnd = Get-Command wt -EA SilentlyContinue
if ( -NOT $wtFnd ) {
    Write-Verbose "Installing Windows Terminal."
    ## install pre-req's ##
    Write-Verbose "Installing terminal pre-requesites."
    # the pre-req's are packages as part of Windows Terminal
    try {
        # get the pre-req url
        Write-Verbose "Get the URL to the pre-req file."
        $preReq = Find-GitRelease $termRepo -Latest -EA Stop
        $prURL = $preReq.URL.split(" ").Trim(" ") | Where-Object { $_ -match $termPRHint }
        Write-Verbose "URL: $prURL"

        # download the file
        $prFile = Get-WebFile -URI $prURL -savePath $savePath -fileName $termPRFile -EA Stop
        Write-Verbose "Downloaded to: $prFile"

        # unzip
        Expand-Archive $prFile -DestinationPath "$savePath\term-pr" -Force -EA Stop

        # install any x64 appx package in the pre-req kit
        Write-Verbose "Installing XAML."
        [array]$appxFiles = Get-ChildItem "$savePath\term-pr" -Filter "*x64*.appx"
        foreach ($ax in $appxFiles) {
            Add-AppxPackage $ax -EA Stop
        }
    } catch {
        return (Write-Error "Pre-requrisite install failed: $_" -EA Stop)
    }

    ## install Windows Terminal ##
    Write-Verbose "Getting Windows Terminal URL."
    $termBundles = Find-GitRelease $termRepo | Where-Object { $_.Name -notmatch "Preview"} | Sort-Object -Property Version -Descending
    $termURL = $termBundles[0].URL.Split(" ").Trim(" ") | Where-Object {$_ -match '^*.msixbundle$'}

    try 
    {
        Write-Verbose "Installing Windows Terminal."
        $termFile = Get-WebFile -URI $termURL -savePath $savePath -fileName "terminal.msixbundle" -EA Stop
        Add-AppxPackage "$termFile" -EA Stop
    }
    catch 
    {
        return (Write-Error "Failed to download or install Windows Terminal: $_" -EA Stop)
    }
}

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (4)" }

# install winget if it is not installed
$wingetFnd = Get-Command winget -EA SilentlyContinue

if ( -NOT $wingetFnd ) {
    Write-Verbose "Installing winget."
    ## download and install winget from github ##
    try 
    {
        # install VCLibs firs
        $libFile = Get-WebFile -URI $vclibsURL -savePath $savePath -fileName $vclibsFile -EA Stop
        Add-AppxPackage "$libFile" -EA Stop
        
        Write-Verbose "Check GitHub for latest winget package."
        $release = Find-GitRelease $wingetRepo -Latest
        $fileName = "$(($wingetRepo -split '/')[-1])`.$wingetExt"

        # find the URL
        $URL = $release.URL.Split(" ").Trim(" ") | Where-Object { $_ -match "^.*.$wingetExt$" }
        Write-Verbose "Winget URL: $URL"

        $installFile = Get-WebFile -URI $URL -savePath $savePath -fileName $fileName -EA Stop
        Write-Verbose "Winget installer saved to: $installFile"

        Write-Verbose "Get winget license file."
        $URL2 = $release.URL.Split(" ").Trim(" ") | Where-Object { $_ -match "^.*.xml$" }
        $fileName2 = Split-Path $url2 -Leaf

        $licenseFile = Get-WebFile -URI $URL2 -savePath $savePath -fileName $fileName2 -EA Stop

        Write-Verbose "Install winget."
        Add-AppxProvisionedPackage -Online -PackagePath $installFile -LicensePath $licenseFile -Verbose -EA Stop
    }
    catch
    {
        return (Write-Error "Winget download failed: $_" -EA Stop)
    }

    # wait for winget to appear in the path
    Write-Verbose "Wait for winget to show up."
    $count = 0
    do
    {
        Start-Sleep 1
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 
        $wingetFnd = Get-Command winget -EA SilentlyContinue
        
        # 2022-09-23 - Using Add-AppxPackage seems to work with winget on WS2022, so try that is Add-AppxProvisionedPackage didn't see to work
        if ($count -gt 0 -and -NOT $wingetFnd)
        {
            Write-Verbose "Attempting a second install of winget"
            Add-AppxPackage $installFile
        }
        
        $count++
    } until ($wingetFnd -or $count -ge 10)
}

## install winget apps ##
if ($wingetFnd)
{
    Write-Verbose "Winget is installed."
    Write-Verbose "Installing apps through winget."
    foreach ($app in $wingetApps)
    {
        Write-Verbose "Installing $app."
        # install things
        winget install $app --exact --accept-package-agreements --accept-source-agreements --silent
    }

    # pwsh doesn't always install the first time. test and retry
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 
    $isPwshFnd = Get-Command pwsh -EA SilentlyContinue

    if (-NOT $isPwshFnd)
    {
        Write-Verbose "Attempting a winget install of PowerShell 7."
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

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (5)" }

## configure PowerShell on Windows Terminal ##
# get CaskaydiaCove NF if not installed
if ($fontName -notin (Get-InstalledFonts))
{
    Write-Verbose "Installing $fontName"
    # get newest font
    $ccnf = Find-GitRelease -repo $repoCCNF -Latest

    # find the correct URL
    $ccnfURL = $ccnf.URL.Split(" ").Trim(" ") | Where-Object {$_ -match $fontFile}

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

$strCMD = @"
    `$nugetVer = Get-PackageProvider -ListAvailable -EA SilentlyContinue | Where-Object Name -match "NuGet" | ForEach-Object { `$_.Version }
    [version]`$minNugetVer = "2.8.5.208"
    if (`$nugetVer -lt `$minNugetVer -or `$null -eq `$nugetVer)
    {
        Write-Verbose "Installing NuGet update."
        `$null = Install-PackageProvider -Name NuGet -MinimumVersion `$minNugetVer -Force
    }

    # get module(s)
    Install-Module -Name "$($pwshMods -join '","')" -Repository PSGallery -Scope CurrentUser -Force
"@

[scriptblock]$cmd = [scriptblock]::Create($strCMD)

# update $env:Path so pwsh will be found
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 

# install pwsh modules
Write-Verbose "Installing pwsh modules."
pwsh -NoLogo -NoProfile -Command $cmd

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (7)" }

# install WSL2 
if ($WSL.IsPresent) {
    Write-Verbose "Installing WSL2."
    $wslCmd = { wsl --install; exit }
    Start-Process pwsh -ArgumentList "-NoLogo -NoProfile -Command $wslCmd" -Wait
}




# update the pwsh profile
Write-Verbose "Updating the PowerShell profile for Terminal."
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
Write-Verbose "Launching pwsh in Terminal to make sure the profile has been created."
Start-Process wt -ArgumentList "-p PowerShell" -WindowStyle Minimized
Start-Sleep 10
# hopefull this closes the correct terminal...
$currPwsh = Get-Process -Id $PID
Get-Process WindowsTerminal | Where-Object { $_.Id -ne $currPwsh.Parent.Id } | Stop-Process -Force

Write-Verbose "Get the Terminal config file."
$appxPack = Get-AppxPackage -Name "Microsoft.WindowsTerminal" -EA SilentlyContinue
$wtAppData = "$ENV:LOCALAPPDATA\Packages\$($appxPack.PackageFamilyName)\LocalState"

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (8)" }

# export settings.json
# clean up comment lines to prevent issues with older JSON parsers (looking at you Windows PowerShell)
try 
{
    $wtJSON =  Get-Content "$wtAppData\settings.json" -EA Stop | Where-Object { $_ -notmatch "^.*//.*$" -and -NOT [string]::IsNullOrEmpty($_)} | ConvertFrom-Json  
    Write-Verbose "Terminal settings have been imported."  
}
catch 
{
    return (Write-Error "Failed to update Windows Terminal settings." -EA Stop)
}

# change the font for PowerShell
if ($null -ne $wtJSON.profiles.list.font.face)
{
    Write-Verbose "Changing pwsh font (1)."
    $wtJSON.profiles.list | Where-Object { $_.Name -eq "PowerShell" } | ForEach-Object { $_.Font.Face = $fontName }
}
else 
{
    Write-Verbose "Changing pwsh font (2)."
    $pwshProfile = $wtJSON.profiles.list | Where-Object { $_.Name -eq "PowerShell" }
    $pwshProfile | Add-Member -NotePropertyName font -NotePropertyValue ([PSCustomObject]@{face="$fontName"})
}

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (9)" }

# set PowerShell (pwsh) to the default profile
Write-Verbose "Set pwsh as the default Terminal profile."
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
Write-Verbose "Disable annoying warnings."
$evilPasteSettings = 'largePasteWarning', 'multiLinePasteWarning'

foreach ($imp in $evilPasteSettings)
{
    if ($null -eq $wtJSON."$imp" -or $wtJSON."$imp" -eq $true)
    {
        $wtJSON | Add-Member -NotePropertyName $imp -NotePropertyValue $false -Force
    }
}

# maximize wt on start
Write-Verbose "Maximize Terminal on start."
$wtJSON | Add-Member -NotePropertyName "launchMode" -NotePropertyValue "maximized" -Force


if ($gDemo.IsPresent) {
    if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (10)" }
    Write-Verbose "Setting demo values."
    
    # add Git Bash to the Terminal settings, if not there
    $bashFnd = $wtJSON.profiles.list | Where-Object { $_.Name -eq "Git Bash" }

    if ( -NOT $bashFnd ) {
        Write-Verbose "Adding Git Bash to Terminal."
        $bashGUID = New-Guid
        $bashProfile = [pscustomobject] @{
            guid   = "{$($bashGUID.Guid)}"
            hidden = $false
            name   = "Git Bash"
            source = "Git"
        }

        $wtJSON.profiles.list += $bashProfile
    }

    ## font stuff ##
    # set the font size and weight
    $wtJSON.profiles.list | ForEach-Object { 
        Write-Verbose "Adding demo font values for $($_.name) (10)."

        # update the settings if the font section is already there
        if ($_.font) {
            if ($_.font.face)   { $_.font.face   = $fontName } else { $_.font | Add-Member -NotePropertyName "face" -NotePropertyValue $fontName }
            if ($_.font.size)   { $_.font.size   = $demoFntSz } else { $_.font | Add-Member -NotePropertyName "size" -NotePropertyValue $demoFntSz }
            if ($_.font.weight) { $_.font.weight = $demoFntWeight } else { $_.font | Add-Member -NotePropertyName "weight" -NotePropertyValue $demoFntWeight }
        # add the font section if it is missing
        } else {
            Write-Verbose "Adding demo font values for $($_.name) (2)."
            $demoFont = [PSCustomObject]@{
                face   = $fontName
                size   = $demoFntSz
                weight = $demoFntWeight
            }
            # add the font section
            $_ | Add-Member -NotePropertyName font -NotePropertyValue $demoFont
        }
    }
}

# save settings
Write-Verbose "Saving Terminal settings."
$wtJSON | ConvertTo-Json -Depth 20 | Out-File "$wtAppData\settings.json" -Force -Encoding utf8

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (11)" }

# get npcap
try 
{
    Write-Verbose "Install npcap."
    $npcapFile = Get-WebFile -URI $npcapURL -savePath $savePath -fileName npcap.exe
    Start-Process "$savePath\npcap.exe" -ArgumentList "/winpcap_mode=disabled" -Wait

    $null = Remove-Item $npcapFile -Force
}
catch 
{
    Write-Warning "Failed to download and install npcap. Please manually download and install: $_"
}

if ($gDemo.IsPresent) {
    if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (12)" }
    Write-Verbose "Setting connection profile to Private."
    # Demo: Set all network connections to Private
    Get-NetConnectionProfile | Where-Object NetworkCategory -eq Public | Set-NetConnectionProfile -NetworkCategory Private

    Write-Verbose "Enabling File and Printer Sharing firewall rules."
    # Demo: Enable File and Printer Sharing on the firewall for Private
    Get-NetFirewallRule -DisplayGroup "File and Printer Sharing" | Where-Object { $_.Profile -eq "Private" -or $_.Profile -eq "Any" } | Enable-NetFirewallRule
}


# update and reboot
Install-PackageProvider -Name NuGet -Force
Install-Module -Name PSWindowsUpdate -MinimumVersion 2.2.0 -Force
Get-WindowsUpdate -AcceptAll -Verbose -WindowsUpdate -Install -AutoReboot
