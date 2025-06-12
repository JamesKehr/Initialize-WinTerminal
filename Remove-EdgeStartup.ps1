# Removes annoying Edge startup prompts. Used for lab systems to disable sync and other first run experienace stuff.
#requires -RunAsAdministrator
#requires -Version 5.1

<#

https://learn.microsoft.com/en-us/deployedge/microsoft-edge-browser-policies/hidefirstrunexperience

HideFirstRunExperience     = 1 (Enabled == hide experience)
AutoImportAtFirstRun       = 4 (DisabledAutoImport)
ForceSync                  = 0 (Disabled)
SyncDisabled               = 1 (Enabled == disable sync)
BrowserSignin              = 0 (Disabled)
NonRemovableProfileEnabled = 0 (Disabled)

NewTabPageContentEnabled         = 0 (Disabled)
NewTabPageAllowedBackgroundTypes = 3 (DisableAll)
NewTabPageQuickLinksEnabled      = 0 (Disabled)

#>


# the root of the Edge registry value path
$edgRootReg = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"

# a hashtable for reg name to reg value 
$edgeRegValues = @{
    HideFirstRunExperience           = 1
    AutoImportAtFirstRun             = 4
    ForceSync                        = 0
    SyncDisabled                     = 1
    BrowserSignin                    = 0
    NonRemovableProfileEnable        = 0
    NewTabPageContentEnabled         = 0
    NewTabPageAllowedBackgroundTypes = 3
    NewTabPageQuickLinksEnabled      = 0
}


# all values are DWORD
$edgeRegType = "Dword"

# create the root path
try {
    $null = New-Item -Path $edgRootReg -Force -EA Stop
} catch {
    Write-Warning "The reg path already exists: $_"
}

# add all the registry values
foreach ($val in $edgeRegValues.Keys) {
    try {
        $null = New-ItemProperty -Path $edgRootReg -Name $val -PropertyType $edgeRegType -Value $edgeRegValues["$val"]
    } catch {
        Write-Warning "Failed to add $val`: $_"
    }
}