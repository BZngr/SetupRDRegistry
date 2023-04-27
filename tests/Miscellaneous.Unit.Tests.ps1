Describe 'Get-InProcServer32PathFromKey'{
    BeforeAll{
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Get-InProcServer32PathFromKey
    }
    
    It 'Works with actual registry key'{
        

        $result = Get-InProcServer32PathFromKey "Registry::HKEY_CLASSES_ROOT\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040}\InProcServer32\2.5.2.0"

        $result | Should -Be "Registry::HKEY_CLASSES_ROOT\CLSID\{69E0F697-43F0-3B33-B105-9B8188A6F040}\InProcServer32"
    }

    It 'Works independent of version number' -ForEach @(
        @{Path = "BogusPath\InProcServer32\2.6.6.6" ; Expected = "BogusPath\InProcServer32"}
        @{Path = "BogusPath\InProcServer32\3.0.1.0" ; Expected = "BogusPath\InProcServer32"}
    ){
        
        $result = Get-InProcServer32PathFromKey $path

        $result | Should -Be $expected
    }
}

Describe 'Get-Rd3HKLMPaths Tests' {
    BeforeAll{
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Get-Rd3HKLMPaths
    }
    It 'Returns 3 paths' -ForEach @(
        @{Base = "RegPath::\CLSID\{XXX}"}
        @{Base = "RegPath::\CLSID\{69E0F699-43F0-3B33-B105-9B8188A6F040}"}
    ) {
        #$base = "RegPath::\CLSID\{XXX}"
        $result = Get-Rd3HKLMPaths $base "3.4.5.6"

        $result.Count | Should -Be 3
        #$result[0] | Should -not -Be $null 
        $result[0] | Should -BeLike "*Registry::HKEY_LOCAL_MACHINE*" 
        $result[0] | Should -not -BeLike "*Registry::HKEY_LOCAL_MACHINE*\InProcServer32*" 
        $result[1] | Should -BeLike "*\InProcServer32"
        $result[2] | Should -BeLike "*\InProcServer32\3.*"
    }
}

Describe 'Get-CurrentRDVersionSetup'{
    BeforeAll{
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Get-CurrentRDVersionSetup

    }
    It 'Returns <expected>' -ForEach @(
        @{RD3Keys = @("RD3Keys", "RD3Keys") ; RD2KKeyFound = $True ; Expected = 3}
        @{RD3Keys = @() ; RD2KeyFound = $False ; Expected = 10}
        @{RD3Keys = @() ; RD2KeyFound = $True ; Expected = 2}
        @{RD3Keys = @("RD3Keys") ; RD2KeyFound = $True ; Expected = 10}
    ) {

        Mock -Command Get-HKLMKeysOfInterest -MockWith {$rd3Keys}

        Mock -Command Test-Path -MockWith {$rd2KeyFound}

        $result = Get-CurrentRDVersionSetup

        $result | Should -Be $expected
    }
}
