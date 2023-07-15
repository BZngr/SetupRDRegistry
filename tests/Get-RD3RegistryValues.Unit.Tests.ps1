Describe 'Get-RegistryProperties'{
    It 'Generates Properties Hashtable'{
        
        . $PSScriptRoot\..\src\Switch-RDVersionCode.ps1 New-RegistryKeyModelFromRegistryKey
        
        $theKey = "BOGUS\Path"

        Mock -CommandName Get-ItemPropertyValue -MockWith {
            Write-Verbose "Called Mock -> Get-ItemPropertyValue" -Verbose
            ("The" + $Name)
        }  

        Mock -Command Test-Path -MockWith {$true}

        $regKeyModel = New-RegistryKeyModelFromRegistryKey $theKey

        $regKeyModel.Properties.Class | Should -Be "TheClass"
        $regKeyModel.Properties.Assembly | Should -BeLike "TheAssembly"
        $regKeyModel.Properties.CodeBase | Should -BeLike "TheCodeBase"
        $regKeyModel.Properties.RuntimeVersion | Should -BeLike "TheRuntimeVersion"
    }
}
