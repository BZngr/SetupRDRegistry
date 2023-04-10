Describe 'New-RdRegistryKeyModel tests'{
    It 'Generates correct number of RD3  registry records'{
        
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1
        
        $theKeys = @("BOGUS\Path1", "BOGUS\Path2")

        $bogusRecord = 
        @{
            "KeyPath" = "ExtensionPath"
            "Class" = "TheClass"
            "CodeBase" = "TheCodeBase"
            "Assembly" = "TheAssembly"
            "RuntimeVersion" = "TheRuntimeVersion"
        }

        Mock -CommandName New-RdRegistryKeyModel -MockWith {
            ($bogusRecord)
        }  

        $values = Get-RegistryKeyModels $theKeys

        $values.Count | Should -Be 2
    }
}
