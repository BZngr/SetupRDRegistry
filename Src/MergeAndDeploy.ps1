
[cmdletbinding(SupportsShouldProcess = $True)]

$scriptContent = Get-Content .\src\SetRegistryForRDVersion.ps1
$implContent = Get-Content .\src\SetRegistryForRDVersionImpl.ps1
$deploymentTarget = "C:\Program Files\WindowsPowerShell\Modules\MyScriptModule\SetRegistryForRDVersion.ps1"

$allLines = @()

foreach ($line in $scriptContent){
    if ($line -like "*#Requires -RunAsAdministrator*"){
        $allLines += "#Requires -RunAsAdministrator"
        continue
    }
    if ($line -like "*SetRegistryForRDVersionImpl.ps1"){
        $allLines += ("#" * 32) + " Script starts at line XXX " +  ("#" * 32)
        $allLines += ""
        foreach ($implLine in $implContent){
            $allLines += $implLine
        }
    }
    $allLines += $line
}

$nextVersion = @()
$lineNum = 0
$scriptStartLine = 0
foreach ($line in $allLines){
    $lineNum++
    if ($line -like "*SetRegistryForRDVersionImpl.ps1"){
        $scriptStartLine = $lineNum
        $nextVersion += ("#" * 32) + " Start of Script " +  ("#" * 32)
        continue
    }
    $nextVersion += $line
}

$finalLines = @()
foreach ($line in $nextVersion){
    if ($line -like "*# Script starts at line XXX #*"){
        $modified = $line.Replace("XXX", $scriptStartLine.ToString())
        $finalLines += $modified
        continue
    }
    $finalLines += $line
}


#$finalLines > ".\TestOutput.txt"
$finalLines | Out-File ".\TestOutput.txt" -Confirm  
#$finalLines > $deploymentTarget