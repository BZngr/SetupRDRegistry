Describe 'Convert-MaybeJsonToRegistryKeyModel tests'{
    BeforeAll {        
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 New-RegistryKeyModelsFromFile
    }
    It 'Returns $null when file contains invalid content'{
        
        $values = New-RegistryKeyModelsFromFile '.\tests\Temp2NonsenseRD3Values.txt'

        $values | Should -Be $null
    }

    It 'Returns $null when file has no content'{
        
        $values = New-RegistryKeyModelsFromFile '.\tests\TempEmptyRD3Values.txt'

        $values | Should -Be $null
    }
}
