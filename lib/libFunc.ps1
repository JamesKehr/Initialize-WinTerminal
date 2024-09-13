# common functions

### GET ###
#region

# installs the latest release for a GH repo.
function script:Get-LatestGitHubRelease {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Owner,

        [Parameter(Mandatory)]
        [string]
        $Repo,

        [Parameter(Mandatory)]
        [string]
        $File,

        [Parameter(Mandatory)]
        [string]
        $Path,

        [Parameter()]
        [string]
        $LicenseFile = $null
    )

    # format the download URL
    $dlURI = "https://github.com/$Owner/$Repo/releases/latest/download/$File"

    # download the file
    Write-Verbose "Get-LatestGitHubRelease - Downloading release from: $dlURI; to: $Path\$File"
    $dlFile = Get-WebFile -URI $dlURI -Path $Path -fileName $File

    # grab the license files when requested
    if (-NOT [string]::IsNullOrEmpty($LicenseFile)) {
        $dlURI = "https://github.com/$Owner/$Repo/releases/latest/download/$LicenseFile"
        Write-Verbose "Downloading license file from: $dlURI; to: $Path\$LicenseFile"
        $null = Get-WebFile -URI $dlURI -Path $Path -fileName $LicenseFile
    }

    return $dlFile
}



# Downloads a file from the Internet.
# Returns the full path to the download.
function Get-WebFile
{
    param ( 
        [string]$URI,
        [string]$Path,
        [string]$FileName
    )

    Write-Debug "Get-WebFile - Start."

    # validate path
    if ( -NOT (Test-Path "$Path" -IsValid) ) {
        return (Write-Error "The save path, $Path, is not valid. Error: $_" -EA Stop)
    }

    # create the path if missing
    if ( -NOT (Get-Item "$Path" -EA SilentlyContinue) ) {
        try {
            $null = mkdir "$Path" -Force -EA Stop
        } catch {
            return (Write-Error "The save path, $Path, does not exist and cannot be created. Error: $_" -EA Stop)
        }
        
    }

    # create the full path
    $OutFile = "$Path\$fileName"

    # use curl if it is found in the path
    # options are iwr (Invoke-WebRequest (default)), bits (Start-BitsTransfer), and curl (preferred when found)
    $dlMethods = "iwr", "curl", "bits"
    $dlMethod = "iwr"

    # switch to curl when found
    $curlFnd = Get-Command "curl.exe" -EA SilentlyContinue
    if ($curlFnd) { $dlMethod = "curl" }

    Write-Verbose "Get-WebFile - Attempting download of $URI to $OutFile"

    # did the download work?
    $dlWorked = $false

    # methods tried
    # initialize with curl because if curl is found then we're using it, if it's not found then we shouldn't try it
    $tried = @("curl")

    # loop through
    do {
        switch ($dlMethod) {
            # tracks whether 
            "curl" {
                Write-Verbose "Get-WebFile - Download with curl."

                Push-Location "$Path"
                # download with curl
                # -L = download location
                # -o = output file
                # -s = Silent
                curl.exe -L $URI -o $OutFile -s
                Pop-Location
            }

            "iwr" {
                Write-Verbose "Get-WebFile - Download with Invoke-WebRequest."

                # make sure we don't try to use an insecure SSL/TLS protocol when downloading files
                Write-Debug "Get-WebFile - Disabling unsupported SSL/TLS protocls."
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12, [System.Net.SecurityProtocolType]::Tls13

                # download silently with iwr
                $oldProg = $global:ProgressPreference
                $Global:ProgressPreference = "SilentlyContinue"
                $null = Invoke-WebRequest -Uri $URI -OutFile "$OutFile" -MaximumRedirection 5 -PassThru
                $Global:ProgressPreference = $oldProg
            }

            "bits" {
                Write-Verbose "Get-WebFile - Download with Start-BitsTransfer."
                
                # download silently with iwr
                $oldProg = $global:ProgressPreference
                $Global:ProgressPreference = "SilentlyContinue"
                $null = Start-BitsTransfer -Source $URI -Destination "$OutFile"
                $Global:ProgressPreference = $oldProg
            }

            Default { return (Write-Error "An unknown download method was selected. This should not happen. dlMethod: $_" -EA Stop) }
        }

        # is there a file, any file, then consider this a success
        $dlFnd = Get-Item "$OutFile" -EA SilentlyContinue

        if ( -NOT $dlFnd ) {
            # change download method and try again
            Write-Verbose "Failed to download using $dlMethod."

            if ($tried.Count -lt $dlMethods.Count) {
                if ($dlMethod -notin $tried) {
                    $tried += $dlMethod
                    Write-Verbose "Get-WebFile - Added $dlMethod to tried: $($tried -join ', ')"
                }

                :dl foreach ($dl in $dlMethods) { 
                    if ($dl -notin $tried) { 
                        Write-Verbose "Get-WebFile - Switching to $dl method."
                        $dlMethod = $dl
                        $tried += $dl
                        break dl
                    }
                }
            } else {
                return (Write-Error "The download has failed!" -EA Stop)
            }
        } else {
            # exit the loop
            $dlWorked = $true
        }
    } until ($dlWorked)

    Write-Verbose "Get-WebFile - File downloaded to $OutFile."

    #Add-Log "Downloaded successfully to: $output"
    Write-Debug "Get-WebFile - Returning: $OutFile"
    Write-Debug "Get-WebFile - End."
    return $OutFile
}

#endregion GET


### MISC ###

# returns the [version] of the latest GH repo release.
# Only major.minor.build are returned! The Revision is ignored because file/local versions may not contain the revision number.
function script:Find-GitHubLatestVersion {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Owner,

        [Parameter(Mandatory)]
        [string]
        $Repo
    )

    [System.Uri]$ghURI = "https://api.github.com/repos/$Owner/$Repo/releases/latest"
    Write-Verbose "Find-GitHubLatestVersion - URI: $ghURI"

    # regex pattern that will find the major.minor.build from an api.github.com REST API get request. Build is optional.
    [string]$pattern = 'v?(?<ver>\d{1,5}\.\d{1,3}(?:\.\d{1,6})?)'

    try {
        # Invoke-RestMethod params
        $imrSplat = @{
            Uri             = $ghURI
            Method          = "Get"
            UseBasicParsing = $true
        }

        $latestRaw = Invoke-RestMethod @imrSplat

        Write-Verbose "Find-GitHubLatestVersion - tagName: $($latestRaw.tag_name)"

        if ($latestRaw.tag_name -match $pattern) {
            [version]$latest = $Matches.ver
        }
        
        Write-Verbose "Find-GitHubLatestVersion - latest: $latest"
    
        return $latest
    } catch {
        throw "Failed to find the latest GitHub release for $Repo. Error: $_"
    }
}
