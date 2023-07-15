Describe 'Get-InProcServer32PathFromKey'{
    BeforeAll{
        . $PSScriptRoot\..\src\Switch-RDVersionCode.ps1 Get-InProcServer32PathFromKey
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
        . $PSScriptRoot\..\src\Switch-RDVersionCode.ps1 Get-Rd3HKLMPaths
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

Describe 'New-TogggleModel'{
    BeforeAll{
        . $PSScriptRoot\..\src\Switch-RDVersionCode.ps1 New-ToggleModel

    }
    It 'Returns <expected>' -ForEach @(
        @{RD2Keys = @("2.5.2.0", "3.0.0.0") ; Expected = "3.0.0.0" ; ExpectedRD2 = "2.5.2.0" ; ExpectedRD3 = "3.0.0.0"}
        @{RD2Keys = @() ; Expected = "TBD" ; ExpectedRD2 = "TBD" ; ExpectedRD3 = "TBD"}
        @{RD2Keys = @("2.5.2.0") ; Expected = "2.5.2.0" ; ExpectedRD2 = "2.5.2.0" ; ExpectedRD3 = "TBD"}
        @{RD2Keys = @("3.0.0.0") ; Expected = "3.0.0.0" ; ExpectedRD2 = "TBD" ; ExpectedRD3 = "3.0.0.0"} #only Rd3 exists on the machine
        @{RD2Keys = @("2.5.9.0", "3.0.0.0") ; Expected = "3.0.0.0" ; ExpectedRD2 = "2.5.9.0" ; ExpectedRD3 = "3.0.0.0"}
        @{RD2Keys = @("2.5.9.0", "3.1.0.0") ; Expected = "3.1.0.0" ; ExpectedRD2 = "2.5.9.0" ; ExpectedRD3 = "3.1.0.0"}
        @{RD2Keys = @("2.5.9.0") ; Expected = "2.5.9.0" ; ExpectedRD2 = "2.5.9.0" ; ExpectedRD3 = "TBD"}
        @{RD2Keys = @("3.1.0.0") ; Expected = "3.1.0.0" ; ExpectedRD2 = "TBD" ; ExpectedRD3 = "3.1.0.0"} #only Rd3 exists on the machine
    ) {

        $result = New-ToggleModel $rd2Keys

        $result.Rd2Version | Should -Be $expectedRD2
        $result.Rd3Version | Should -Be $expectedRD3
        $result.ActiveVersion | Should -Be $expected
    }
}

