Describe 'Convert-MaybeJsonToRegistryKeyModel tests'{
    It 'Returns $null when file contains invalid content'{
        
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 New-RegistryKeyModelsFromFile
        
        $values = New-MaybeRegistryKeyModelsFromFile '.\tests\Temp2NonsenseRD3Values.txt'

        $values | Should -Be $null
    }

    It 'Returns $null when file has no content'{
        
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 New-RegistryKeyModelsFromFile
        
        $values = New-MaybeRegistryKeyModelsFromFile '.\tests\TempEmptyRD3Values.txt'

        $values | Should -Be $null
    }
}
