##Requires -RunAsAdministrator

[cmdletbinding(SupportsShouldProcess = $True)]
param (
    [Parameter(Mandatory = $True)]
    [ValidateSet("CurrentVersion", "rd2", "rd3", "2", "3")] 
    [string]$targetMajorVersion
)

. $PSScriptRoot\SetRegistryForRDVersionImpl.ps1

if ($targetMajorVersion -eq "CurrentVersion"){
    switch(Get-CurrentRDVersionSetup)
    {
        2 {"Current Registry configuration: Rubberduck Version 2"}
        3 {"Current Registry configuration: Rubberduck Version 3"}
        Default {"Unknown: Rubberduck2 and/or Rubberduck3 may not be installed"}
    }
    exit
}

$validationResults = Test-CanProcessToTargetVersion ($targetMajorVersion -like "*2")

if (-not $validationResults.CanProcess){
    $validationResults.Message
    exit
}

if ($targetMajorVersion -like "*2"){
    Restore-RubberduckV2 $validationResults.Rd3HKLMClsidKeys

} else {
    Restore-RubberduckV3 $validationResults.ModelsFromFile
}

<#
.SYNOPSIS
Modifies the registry to enable either the Rubberduck 3 or the Rubberduck 2.5.2.0 VBIDE Add-In.
.DESCRIPTION
Modifies the registry to enable either the Rubberduck 3 or the Rubberduck 2.5.2.0 VBIDE Add-In.

Why: 
RD3 re-uses RD2 CLSIDs.  Consequently, the two Add-Ins cannot co-exist in memory.  
When the RD3 solution is built, the build process introduces new keys to the LocalMachine (HKLM) hive.  
Re-building RD2 does not configure the registry to load RD2 once the HKLM keys have been introduced.

*** The script must be invoked from a PowerShell session with Administrator privileges ***

To run the script to see what it does WITHOUT making changes, enter the following:
(The commands below assume a Powershell session is opened in the folder containing the SetRegistryForRDVersion.ps1 script)

Setup for RD2:    PS> .\SetRegistryForRDVersion.ps1 2 -Verbose -WhatIf
Setup for RD3:    PS> .\SetRegistryForRDVersion.ps1 3 -Verbose -WhatIf

****************************************************************************************
Note: Absence of the -WhatIf switch parameter from the above expressions will enable the 
script to make changes to the registry.
*****************************************************************************************

To see the currently active version: 
PS> .\SetRegistryForRDVersion.ps1 CurrentVersion

Alternatively, re-building the RD3 solution will also modify the Registry, enabling the RD3 Add-In  

There is more content available using the -examples, -detailed, or -full switches as indicated in REMARKS
.NOTES
Setting up for Rubberduck Version 2:

1. Deletes two HKLM registry keys (including their children) introduced by building RD3: 
    HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040} (Rubberduck.Extension)
    HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{69E0F699-43F0-3B33-B105-9B8188A6F040} (Rubberuck.DockableToolWindow)
    Note: before deleting, the script saves values contained in these keys to a file (RD3HKLMValues.json) located in 
        the same folder as the Rubberduck.config file.

2. Modifies the values of:
    HKEY_CLASSES_ROOT\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040}\InProcServer32 and
    HKEY_CLASSES_ROOT\CLSID\{69E0F699-43F0-3B33-B105-9B8188A6F040}\InProcServer32

    to match the values contained in 

    HKEY_CLASSES_ROOT\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040}\InProcServer32\2.5.2.0 and 
    HKEY_CLASSES_ROOT\CLSID\{69E0F699-43F0-3B33-B105-9B8188A6F040}\InProcServer32\2.5.2.0 respectively

Setting up for Rubberduck Version 3:

1. Re-creates the RD3 HKLM registry keys and their children based on data contained in 
    RD3HKLMValues.json.  The RD3HKLMValues.json file is generated/refreshed when the script is run 
    to update from RD3 to RD2.

2. Modifies the following registry keys with RD3 values

    HKEY_CLASSES_ROOT\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040}\InProcServer32 and
    HKEY_CLASSES_ROOT\CLSID\{69E0F699-43F0-3B33-B105-9B8188A6F040}\InProcServer32


3. If the RD3HKLMValues.json file is missing or cannot be used for any reason, the user is prompted to 
rebuild the Rubbberduck3 solution.  Rebuilding the solution results in the same registry changes
as described in steps 1 and 2 above. 

Note: Closing and re-opening the Registry Editor application will show that adding the two HKLM keys 
in #1 above, resulted in the creation of:
    HKEY_CLASSES_ROOT\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040}\InProcServer32\3.X.X.X and 
    HKEY_CLASSES_ROOT\CLSID\{69E0F699-43F0-3B33-B105-9B8188A6F040}\InProcServer32\3.X.X.X

.EXAMPLE
...\SetRegistryForRDVersion.ps1 CurrentVersion
Script output identifies the Rubberduck version currently setup in the registry

.EXAMPLE
...\SetRegistryForRDVersion.ps1 2 -WhatIf
See operations that 'would have' taken place to setup for RD2

.EXAMPLE
...\SetRegistryForRDVersion.ps1 2 -Verbose -WhatIf
See operations that 'would have' taken place to setup for RD2 with the most descriptive content

.EXAMPLE
...\SetRegistryForRDVersion.ps1 rd2 -Verbose
Allow the script to make changes to the registry to setup RD2, providing a narrative of the process

.EXAMPLE
.\SetRegistryForRDVersion.ps1 3 -Verbose -WhatIf
See operations that 'would have' taken place to setup for RD3 with the most descriptive content

.EXAMPLE
...\SetRegistryForRDVersion.ps1 rd3 -Verbose -WhatIf
See operations that 'would have' taken place to setup for RD3 with the most descriptive content

.EXAMPLE
...\SetRegistryForRDVersion.ps1 rd3
Allow the script to make changes to the registry to setup RD3 with limited descriptive content
#>


