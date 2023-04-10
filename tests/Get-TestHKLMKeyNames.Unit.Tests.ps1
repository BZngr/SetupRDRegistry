Describe "Get-TestHKLMKeyNames Tests"{
    BeforeAll{
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Get-HKLMKeysOfInterest

        function Get-RDCLSIDs(){
            @(
                "{69E0F697-43F0-3B33-B105-9B8188A6F040}"
                "{69E0F699-43F0-3B33-B105-9B8188A6F040}"
            )
        }
        function Get-NonRDCLSIDs(){
            @(
                "{69E0F697-43F0-3B33-B105-9B8188A6F0XA}"
                "{69E0F699-43F0-3B33-B105-9B8188A6F0XB}"
                "{69E0F697-43F0-3B33-B105-9B8188A6F0XC}"
                "{69E0F699-43F0-3B33-B105-9B8188A6F0XD}"
                "{69E0F611-43F0-3B33-B105-9B8188A6F0XC}"
                "{69E0F612-43F0-3B33-B105-9B8188A6F0XD}"
            )
        }

        function Get-TestHKLMKeyNames(){
            (Get-RDCLSIDs) + (Get-NonRDCLSIDs)
        }
    }
    It "Should find 2 HKLM paths"{     
        $result = Get-HKLMKeysOfInterest (Get-TestHKLMKeyNames)

        $result.Count | Should -Be 2
    }

    It "Should find the correct CLSID Keys"{
        $expectedWildcard = "Registry::HKey_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{69E0F6*-43F0*"
        $result = Get-HKLMKeysOfInterest (Get-TestHKLMKeyNames)

        $result[0] | Should -BeLike $expectedWildcard
        $result[1] | Should -BeLike $expectedWildcard
    }
}