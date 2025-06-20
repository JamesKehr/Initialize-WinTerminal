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
        $installFile = Get-WebFile -URI $URL -Path $savePath -FileName $fileName -EA Stop
        
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
    # load the assembly 
    if ( -NOT ([appdomain]::currentdomain.getassemblies() | Where-Object {$_ -match "System.Drawing"})) {
        Add-Type -AssemblyName System.Drawing
    }

    return ((New-Object System.Drawing.Text.InstalledFontCollection).Families.Name)
}



function Install-Font {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Path
    )

    Write-Verbose "Install-Font - Path: $Path"
    # load the assembly 
    if ( -NOT ([appdomain]::currentdomain.getassemblies() | Where-Object {$_ -match "System.Drawing"})) {
        Add-Type -AssemblyName System.Drawing
    }

    # font directory
    #$fontDir = [System.Environment]::GetFolderPath("Fonts")
    #$fontReg = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'

    # create a font collection to read the font family name later on
    $FontCollection = [System.Drawing.Text.PrivateFontCollection]::new()

    # get a list of installed font families
    #$installedFonts = Get-InstalledFonts

    # get a list of installed ttf and otf fonts
    #$installedFontFiles = Get-ChildItem -Path $fontDir -Include *.otf, *.ttf -File
    
    # stuff needed for CopyHere method
    $FONTS = 0x14
    $CopyOptions = 4 + 16 + 1024
    $objShell = New-Object -ComObject Shell.Application
    $objFolder = $objShell.Namespace($FONTS)

    # get the fonts from the path
    # only use static TTF fonts
    [System.Collections.Generic.List[Object]]$fonts = Get-ChildItem -Path "$Path\ttf\" -Filter "*.ttf" -File
    Write-Verbose "Install-Font - Fonts:`n$($fonts -join "`n")"

    <#Get-ChildItem -Path $Path -Include  -File | & {process {
        if ($_.BaseName -notin $fonts.BaseName) {
            $fonts.Add($_)
        }
    }}#>

    foreach ($font in $fonts) {
        # get font details
        $null = $FontCollection.AddFontFile("$($font.fullname)")
        $currentFont = $FontCollection.Families[-1].Name
        Write-Verbose "Processing: $currentFont"
        #Write-Verbose "FontCollection:$($c=0; $FontCollection.Families | % {"`n$c`t: $($_.Name)"; $c++})"
        
        <# use a standard copy to update the existing font if it already exists
        if ($currentFont -in $installedFonts -or $font.Name -in $installedFontFiles.Name) {
            Write-Verbose "Updating font."
            $null = Copy-Item -Path "$($font.FullName)" -Destination $fontDir -Force -Confirm:$false
        # use CopyHere if the font is not found
        } else {#>
            Write-Verbose "Installing font."
            $CopyFlag = [String]::Format("{0:x}", $CopyOptions);
            $objFolder.CopyHere($font.fullname,$CopyFlag)
        #}
    }
}

#endregion FUNCTIONS

### CONSTANTS ###
#region

# load the common functions
. "$PSScriptRoot\lib\libFunc.ps1"

# where to put downloads
$savePath = "C:\Temp"

# list of exact winget app IDs to install
[array]$wingetApps = "Microsoft.PowerShell", 
                        "Microsoft.WindowsTerminal", 
                        "Microsoft.VisualStudioCode", 
                        "JanDeDobbeleer.OhMyPosh", 
                        "WiresharkFoundation.Wireshark", 
                        "Git.Git"

# powershell repro and file extension, installed by winget now
#$pwshRepoOwner = "PowerShell"
#$pwshRepo = "PowerShell"
#$pwshExt = "msi"

# Windows Terminal details - installed by winget
#$termRepoOwner = "Microsoft"
#$termRepo = "Terminal"
#$termPRHint = "PreinstallKit"
#$termPRFile = "term-pr.zip"

# VCLibs Desktop, XAML, and winget are installed by Update-WingetApps

# repro for Caskaydia Cove Nerd Font
$repoCCNFOwner = "microsoft"
$repoCCNFRepo = "cascadia-code"

# name of the preferred pretty font, CaskaydiaCove NF
$fontName = "Cascadia Code NF"

# the zip file where CC NF is in
$fontVer = Find-GitHubLatestVersion -Owner $repoCCNFOwner -Repo $repoCCNFRepo
$fontFile = "CascadiaCode-$($fontVer.Major).$($fontVer.Minor).zip"

# npcap URL
# Wireshark required npcap
$npcapURL = 'https://npcap.com/dist/npcap-1.80.exe'



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
                              'oh-my-posh --init --shell pwsh --config $env:POSH_THEMES_PATH/rudolfs-dark.omp.json | Invoke-Expression',
                              'cls'

    # PowerShell modules to install                              
    $pwshMods = "Terminal-Icons", "oh-my-posh"
}


# list of apps to remove
$removeApps =   "Clipchamp.Clipchamp",
                "Microsoft.BingNews",
                "Microsoft.BingWeather",
                "Microsoft.Edge.GameAssist",
                "Microsoft.GamingApp",
                "Microsoft.GetHelp",
                "Microsoft.MicrosoftOfficeHub",
                "Microsoft.MicrosoftSolitaireCollection",
                "Microsoft.MicrosoftStickyNotes",
                "Microsoft.People",
                "Microsoft.PowerAutomateDesktop",
                "Microsoft.Todos",
                "Microsoft.WindowsCamera",
                "microsoft.windowscommunicationsapps",
                "Microsoft.WindowsMaps",
                "Microsoft.WindowsSoundRecorder",
                "crosoft.Xbox.TCUI",
                "Microsoft.XboxGameOverlay",
                "Microsoft.XboxGamingOverlay",
                "Microsoft.XboxSpeechToTextOverlay",
                "Microsoft.YourPhone",
                "Microsoft.ZeVideo",
                "Microsoft.Getarted",
                "Microsoft.XboxIdentityProvider",
                "Microsoft.Copilot",
                "Microsoft.ZuneMic"

#endregion CONSTANTS
### MAIN ###",

<# only works for Windows right no
ite-Error "Only Windows is supported at the moment." -EA Stop)
}#>

# prompt for computer name
$compy = Read-Host "Enter the computer name or input nothing and press Enter to bypass computer rename"
if (-NOT [string]::IsNullOrEmpty($compy) -and -NOT [string]::IsNullOrWhiteSpace($compy)) {
    Rename-Computer -NewName $compy -Force
}

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

# install or update winget and its dependencies
try {
    Write-Verbose "Installing winget."
    Push-Location "$PSScriptRoot" 
    .\Update-WingetApps.ps1 -Path $savePath -EA Stop
} catch {  
    throw "Failed to install Winget: $_"
} finally {
    Pop-Location
}

## install winget apps ##
$wingetFnd = Get-Command winget -EA SilentlyContinue
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
    # download
    try {
        $ccnfZip = Get-LatestGitHubRelease -Owner $repoCCNFOwner -Repo $repoCCNFRepo -File $fontFile -Path $savePath -EA Stop
        
        # extract
        $extractPath = "$savePath\ccnf"
        Expand-Archive -Path $ccnfZip -DestinationPath $extractPath -Force -EA Stop

        # install fonts
        Install-Font "$extractPath" -EA SilentlyContinue

        #Start-Sleep 10
    } catch {
        Write-Error "Failed to download $fontFile. Please download and install $fontName manually, or the Nerd Font of your choice."
    }
}

## install modules ##

$strCMD = @"
    # install module(s)
    Install-PSResource -Name "$($pwshMods -join '","')" -Repository PSGallery -Scope CurrentUser -TrustRepository
"@

# do not encode the command as it gives some endpoint protect services a conniption fit
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

Write-Verbose "pwshProfile: $pwshProfile"

if (-NOT (Test-path $pwshProfilePath)) { $null = New-Item $pwshProfilePath -ItemType File -Force }

foreach ($line in $profileLines)
{
    if ($line -notin $pwshProfile)
    {
        $line | Out-File "$pwshProfilePath" -Append -Force
    }
}

# assume WT is installed at this point
# only do this if WT is not currently running
$wtRunning = Get-Process WindowsTerminal -EA SilentlyContinue
if ( -NOT $wtRunning ) {
    # launch WT once to make sure settings.json is generated
    Write-Verbose "Launching pwsh in Terminal to make sure the profile has been created."
    Start-Process wt -ArgumentList "-p PowerShell" -WindowStyle Minimized
    Start-Sleep 10
    Get-Process WindowsTerminal | Stop-Process -Force
} else {
    Write-Verbose "Windows Terminal is running. Opening a new window to generate the PowerShell profile."
    [array]$currWT = Get-Process WindowsTerminal
    wt --profile PowerShell
    Start-Sleep 10
    [array]$newWT = Get-Process WindowsTerminal
    $stopWT = Compare-Object $currWT.Id $newWT.Id -PassThru | ForEach-Object {$_}
    Stop-Process -Id $stopWT
}

Write-Verbose "Get the Terminal config file."
$appxPack = Get-AppxPackage -Name "Microsoft.WindowsTerminal" -EA SilentlyContinue
$wtAppData = "$ENV:LOCALAPPDATA\Packages\$($appxPack.PackageFamilyName)\LocalState"

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (8)" }

# export settings.json
# clean up comment lines to prevent issues with older JSON parsers (looking at you Windows PowerShell)
try {
    $wtJSON =  Get-Content "$wtAppData\settings.json" -EA Stop | Where-Object { $_ -notmatch "^.*//.*$" -and -NOT [string]::IsNullOrEmpty($_)} | ConvertFrom-Json  
    Write-Verbose "Terminal settings have been imported."  
} catch {
    return (Write-Error "Failed to update Windows Terminal settings." -EA Stop)
}

# change the font for PowerShell
if ($null -ne $wtJSON.profiles.list.font.face) {
    Write-Verbose "Changing pwsh font (1)."
    $wtJSON.profiles.list | Where-Object { $_.Name -eq "PowerShell" } | ForEach-Object { $_.Font.Face = $fontName }
} else {
    Write-Verbose "Changing pwsh font (2)."
    $pwshProfile = $wtJSON.profiles.list | Where-Object { $_.Name -eq "PowerShell" }
    if ($pwshProfile) {
        $pwshProfile | Add-Member -NotePropertyName font -NotePropertyValue ([PSCustomObject]@{face="$fontName"})
    }
    
}

if ($gDemo.IsPresent) { Write-Debug "Demo mode engaged. (9)" }

# set PowerShell (pwsh) to the default profile
Write-Verbose "Set pwsh as the default Terminal profile."
$pwshGUID = $wtJSON.profiles.list | Where-Object Name -eq "PowerShell" | ForEach-Object { $_.guid }

if ($pwshGUID) {
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
    $npcapFile = Get-WebFile -URI $npcapURL -Path $savePath -FileName npcap.exe
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

# remove all the Edge startup nonsense
Push-Location "$PSScriptRoot"
.\Remove-EdgeStartup.ps1
Pop-Location


# remove apps
$allApps = Get-AppxPackage
$allUsersApps = Get-AppxPackage -AllUsers

foreach ($app in $removeApps) {
    if ($app -in $allApps.Name) {
        Get-AppxPackage $app | Remove-AppxPackage -Confirm:$false
    }

    if ($app -in $allUsersApps.Name) {
        Get-AppxPackage $app -AllUsers | Remove-AppxPackage -Confirm:$false -AllUsers
    }
}

# update and reboot
Install-PackageProvider -Name NuGet -Force
Install-Module -Name PSWindowsUpdate -MinimumVersion 2.2.1.4 -Force
Get-WindowsUpdate -AcceptAll -Verbose -WindowsUpdate -Install -AutoReboot
