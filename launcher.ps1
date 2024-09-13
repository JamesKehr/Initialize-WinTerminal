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

# find the script path, use PWD as the backup
if ( [string]::IsNullOrEmpty($PSScriptRoot) ) {
    # set $PWD to scriptPath
    $script:scriptPath = $PWD.Path
} else {
    $script:scriptPath = $PSScriptRoot
}

Write-Verbose "scriptPath: $scriptPath"

# make sure all the required files are present and try to download anything missing
[array]$reqFiles = Invoke-WebRequest 'https://raw.githubusercontent.com/JamesKehr/Initialize-WinTerminal/main/file.json' | ForEach-Object Content | ConvertFrom-Json

# loop through each required file
foreach ($rf in $reqFiles) {
    # convert to hashtable
    $htRF = $rf.PSObject.Properties | ForEach-Object -Begin {$h = @{}} -Process {$h."$($_.Name)" = $_.Value} -End {$h}

    # replace . in the path with $scriptPath
    $htRF.Path = $htRF.Path -replace "\.", "$scriptPath"

    if ( -NOT (Get-Item "$($htRF.Path)\$($htRF.FileName)" -EA SilentlyContinue) ) {
        Write-Verbose "Downloading a missing file. htRF:`n$($htRF | Format-List | Out-String)`n"
        
        # download
        $tryDL = $fileList = Get-WebFile @htRF
        if (-NOT $tryDL) {
            throw "Failed to download a required file."
        }
    } else {
        Write-Verbose "File, $($htRF.Path)\$($htRF.FileName), was found."
    }
}

# launch Initialize-WinTerminal.ps1
Start-Process powershell -ArgumentList "-NoLogo -NoProfile -File .\Initialize-WinTerminal.ps1" -WorkingDirectory "$scriptPath"
