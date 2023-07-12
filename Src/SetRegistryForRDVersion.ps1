##Requires -RunAsAdministrator

[cmdletbinding(SupportsShouldProcess = $True)]
param (
    [Parameter(Mandatory = $False)]
    [ValidateSet("CurrentVersion", "rd2", "rd3", "2", "3")] 
    [string]$requestedOp
)

. $PSScriptRoot\SetRegistryForRDVersionImpl.ps1

function ShowUI(){
    param (
        [Parameter(Mandatory = $True)]
        [hashtable]$viewModel
    )

    $theXaml = 
@"
<?xml version="1.0" encoding="utf-8"?><Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
Title="Rubberduck Version Toggling Script" Height="450" Width="600" SizeToContent="WidthAndHeight">
<Window.Resources><Style TargetType="Button"><Setter Property="Padding" Value="10, 5, 10, 5" /><Setter Property="FontSize" Value="20" /><Setter Property="VerticalAlignment" Value="Top" /><Setter Property="FontWeight" Value="Bold" /></Style></Window.Resources>
<DockPanel HorizontalAlignment="Center"><StackPanel HorizontalAlignment="Center" DockPanel.Dock="Top">
<StackPanel HorizontalAlignment="Center" Margin="20,20,20,30"><TextBlock FontSize="20.0" HorizontalAlignment="Center"><Bold>Toggle Rubberduck Registry Setup</Bold></TextBlock><Button Name="KeepCurrent" Margin="10,15,3,10" HorizontalAlignment="Center"><StackPanel><TextBlock FontSize="12.0" HorizontalAlignment="Center"><Bold>(Current Setup)</Bold></TextBlock><TextBlock Name="CurrentVButtonLabel" FontSize="20.0" HorizontalAlignment="Center"><Bold>v2.5.9.x</Bold></TextBlock></StackPanel><Button.ToolTip><TextBlock Margin="5,0,5,0" Width="300" TextWrapping="WrapWithOverflow"><Bold><Italic>DOES NOT modify the Registry</Italic></Bold><LineBreak /><LineBreak />&#xD;
Dismisses the dialog box&#xD;
</TextBlock></Button.ToolTip></Button><StackPanel Name="ShowToggleButtons"><Button Name="SwitchVersion" Margin="10,15,3,0" HorizontalAlignment="Center"><StackPanel><TextBlock FontSize="12.0" HorizontalAlignment="Center"><Bold>(Switch to)</Bold></TextBlock><TextBlock Name="SwitchToVButtonLabel" FontSize="16.0" HorizontalAlignment="Center"><Bold>v3.0.0.x</Bold></TextBlock></StackPanel><Button.ToolTip><StackPanel><TextBlock Margin="5,0,5,0" Width="300" TextWrapping="WrapWithOverflow"><Bold><Italic>MODIFIES THE REGISTRY</Italic></Bold><LineBreak /><LineBreak />Enables the indicated version&#xD;
<LineBreak /></TextBlock></StackPanel></Button.ToolTip></Button><Button Name="Preview" Margin="10,40,3,0" HorizontalAlignment="Center" ToolTipService.ShowDuration="15000"><StackPanel><TextBlock FontSize="12.0" HorizontalAlignment="Center"><Italic><Bold>Preview Registry changes for</Bold></Italic></TextBlock><TextBlock Name="PreviewVButtonLabel" FontSize="12.0" HorizontalAlignment="Center"><Bold>v3.0.0.x</Bold></TextBlock></StackPanel><Button.ToolTip><TextBlock Margin="5,0,5,0" Width="300" TextWrapping="WrapWithOverflow"><Bold><Italic>DOES NOT modify the Registry</Italic></Bold><LineBreak /><LineBreak />&#xD;
Posts messages to the Powershell &#xD;
session window providing a detailed narrative of changes &#xD;
that <Italic>would have</Italic> been made. &#xD;
<LineBreak /><LineBreak />  &#xD;
Executes the script declaring both the '-Verbose' and '-WhatIf' &#xD;
switch parameters.&#xD;
</TextBlock></Button.ToolTip></Button></StackPanel><StackPanel Name="ShowError" Visibility="Collapsed"><TextBlock Name="ErrorMessage" Margin="0, 10, 0, 0" HorizontalAlignment="Center" FontSize="20" TextWrapping="WrapWithOverflow" Foreground="Red"><Bold>This is the error message</Bold><LineBreak /></TextBlock><TextBlock Name="FixMessage" HorizontalAlignment="Center" TextWrapping="WrapWithOverflow"><Bold>This is how you fix it!!</Bold></TextBlock></StackPanel></StackPanel></StackPanel><Border Background="Gray" BorderBrush="SteelBlue" BorderThickness="3,5,3,5"><StackPanel><Button Name="Cancel" IsCancel="True" FontSize="12" HorizontalAlignment="Right" DockPanel.Dock="Bottom" Padding="10, 5, 10, 5" Margin="2,5,10,5">Exit&#xD;
</Button></StackPanel></Border></DockPanel></Window>
"@    
    
    Add-Type -Assembly PresentationFrameWork

    trap {
        $viewModel.Action = "Cancel" 
        $viewModel 
        exit
    }

    $mode = 
        [System.Threading.Thread]::CurrentThread.ApartmentState
        if ($mode -ne "STA"){
            $m = "This  script can only be run when PS is started with the -sta switch."
            throw $m
        }

    $viewModel.Action = "Cancel" 

    $uiWindow = New-WindowFromXamlBlob $theXaml

    Set-UIButtonVersionLabels $uiWindow $viewModel

    $viewModel.Preview = $false

    $keepCurrent = $uiWindow.FindName("KeepCurrent")
    $keepCurrent.add_Click({
        $viewModel.Action = "Cancel" 
        $uiWindow.Close()
    })

    $dismiss = $uiWindow.FindName("Cancel")
    $dismiss.add_Click({
        $viewModel.Action = "Cancel" 
        $uiWindow.Close()
    })

    if ($viewModel.HasDisplayError){
        $btnPanel = $uiWindow.FindName("ShowToggleButtons")
        $btnPanel.Visibility = "Collapsed"

        $errPanel = $uiWindow.FindName("ShowError")
        $errPanel.Visibility = "Visible"

        $errMsg = $uiWindow.FindName("ErrorMessage")
        $errMsg.Text = $viewModel.ErrMsg.Primary
        $theFix = $uiWindow.FindName("FixMessage")
        $theFix.Text = $viewModel.ErrMsg.Fix

    } else {
        $switchVersion = $uiWindow.FindName("SwitchVersion")
        $switchVersion.add_Click({
            $viewModel.Action = "Switch" 
            $viewModel.Preview = $false
            $uiWindow.Close()
        })
    
        $preview = $uiWindow.FindName("Preview")
        $preview.add_Click({
            $viewModel.Action = "Switch" 
            $viewModel.Preview = $true
            $uiWindow.Close()
        })
    }

    if ($uiWindow.ShowDialog()){}

    $viewModel
}

Initialize-ScriptScopeFields

#If the user did not provide CL a parameter, provide the UI
if (-not $PSBoundParameters.ContainsKey('requestedOp')){

    $vm = ShowUI (
        New-ViewModel (Get-ActiveVersion) (Get-InActiveVersion) (Get-Rd2Version) (Get-RD3Version)
    )

    try {
        $VerbosePreference = "Continue"
        $WhatIfPreference = $vm.Preview
    
        if($vm.Action -eq "Cancel"){
            Write-Verbose ("Current Registry Configuration: Rubberduck " + (Convert-VersionToXVersion($vm.CV)))
            $VerbosePreference = "SilentlyContinue"
            $WhatIfPreference = $false
            exit
        }
    
        if (Test-IsConfiguredForRD2){
            Restore-RubberduckV3 $script:rd3ModelsFromFile
        }
    
        if (Test-IsConfiguredForRD3){
            $rd3HKLMKeys = Get-HKLMKeysOfInterest( Get-SubKeyNames "LocalMachine" "SOFTWARE\Classes\CLSID")
            Restore-RubberduckV2 $rd3HKLMKeys
        }
    }
    finally {
        $VerbosePreference = "SilentlyContinue"
        $WhatIfPreference = $false
    }
    exit
}

if ($requestedOp -eq "CurrentVersion"){
    
    $activeVersion = Get-ActiveVersion

    if ($activeVersion -eq "TBD"){ 
        "Unknown: Rubberduck2 and/or Rubberduck3 may not be installed"
    } else {
        "Current Registry configuration: Rubberduck Version " + (Convert-VersionToXVersion $activeVersion)
    }
    exit
}

$targetFromRequestedOpIsRD2 = ($requestedOp -like "*2")

$validationResults = Test-CanProcessToTargetVersion $targetFromRequestedOpIsRD2

if (-not $validationResults.CanProcess){
    $validationResults.Message
    exit
}

if ($targetFromRequestedOpIsRD2){
    $rd3HKLMKeys = Get-HKLMKeysOfInterest (Get-SubKeyNames "LocalMachine" "SOFTWARE\Classes\CLSID")
    Restore-RubberduckV2 $rd3HKLMKeys

} else {
    if ($script:toggleModel.Rd3Version -eq "TBD"){
        $keyPaths = $validationResults.ModelsFromFile.KeyPath
        $script:toggleModel.Rd3Version = Get-Rd3VersionFromRd3RegistryKey($keyPaths[0])
    }
    Restore-RubberduckV3 $validationResults.ModelsFromFile
}

<#
.SYNOPSIS
Modifies the registry to enable either the Rubberduck 3 or the Rubberduck 2 VBIDE Add-In.
.DESCRIPTION
Modifies the registry to enable either the Rubberduck 3 or the Rubberduck 2 VBIDE Add-In.

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

    HKEY_CLASSES_ROOT\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040}\InProcServer32\2.X.Y.Z and 
    HKEY_CLASSES_ROOT\CLSID\{69E0F699-43F0-3B33-B105-9B8188A6F040}\InProcServer32\2.X.Y.Z respectively

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
    HKEY_CLASSES_ROOT\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040}\InProcServer32\3.X.Y.Z and 
    HKEY_CLASSES_ROOT\CLSID\{69E0F699-43F0-3B33-B105-9B8188A6F040}\InProcServer32\3.X.Y.Z

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


