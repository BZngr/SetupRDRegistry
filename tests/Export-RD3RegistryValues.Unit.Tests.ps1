Describe 'Export-RD3RegistryValues'{
    It 'Export/Imports RegistryKeyModels'{
        
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Export-RD3RegistryValues
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 New-RegistryKeyModelsFromFile
                        
        $testDataPath = $PSScriptRoot + "\TempPersistedTestValues.txt"
        if (Test-Path $testDataPath){
            Remove-Item -Path $testDataPath
        }

        $registryRecord1 = @{
            "KeyPath" = "KeyPath1" 
            "Properties" = 
            @{
                "Class" = "TestClass1"
                "RunTimeVersion" = "TestRunTimeVersion1"
                "Assembly" = "TestAssembly1"
                "CodeBase" = "TestCodeBase1"
            }
        }

        $registryRecord2 = @{
            "KeyPath" = "KeyPath2" 
            "Properties" = 
            @{
                "Class" = "TestClass2"
                "RunTimeVersion" = "TestRunTimeVersion2"
                "Assembly" = "TestAssembly2"
                "CodeBase" = "TestCodeBase2"
            }
        }
        
        Export-RD3RegistryValues @($registryRecord1, $registryRecord2) $testDataPath

        Test-Path $testDataPath | Should -Be $true

        $regDataObjects = New-MaybeRegistryKeyModelsFromFile $testDataPath
        
        for ($idx = 0; $idx -lt $regDataObjects.Count; $idx++){
            $numeric = (($idx + 1).ToString())
            $regDataObjects[$idx].KeyPath | Should -Be ("KeyPath" + $numeric)
            $regDataObjects[$idx].Properties.Class | Should -Be ("TestClass" + $numeric)
            $regDataObjects[$idx].Properties.Assembly | Should -Be ("TestAssembly" + $numeric)
            $regDataObjects[$idx].Properties.RuntimeVersion | Should -Be ("TestRuntimeVersion" + $numeric)
            $regDataObjects[$idx].Properties.CodeBase | Should -Be ("TestCodeBase" + $numeric)    
        }
    }
}
