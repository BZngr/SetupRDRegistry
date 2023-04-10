Describe 'Remove-RD3RegistryKeys Calls Remove-Item'{
    It 'Removes 2 HKLM Keys'{
        
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Remove-RD3RegistryKeys

        Mock -CommandName Remove-Item -MockWith {
            Write-Verbose -Message "Called Mock -> Remove-Item" -Verbose
        }  
        
        Mock -Command Test-Path -MockWith {$true}
        
        Remove-RD3RegistryKeys @("Bogus-Key1", "Bogus-Key2")

        Should -Invoke -CommandName Remove-Item -Times 2
    }
}
