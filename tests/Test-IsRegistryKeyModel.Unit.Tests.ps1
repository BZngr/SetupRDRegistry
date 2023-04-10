Describe 'Test-IsValidRegistryKeyModel Tests'{
    BeforeAll {
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Test-IsValidRegistryKeyModel
    }
    It 'All expected Keys are present and not null' -ForEach @(
        @{Model = $null ; Expected = $False}
        @{Model = @{KeyPath = $null ; Properties = @{}} ; Expected = $False}
        @{Model = @{KeyPath = "APath" ; Properties = @{}} ; Expected = $True}
        @{Model = @{KeyPath = "APath" ; Properties = $null} ; Expected = $False}
    ) {
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Test-HasAllProperties
        
        Mock -Command Test-HasAllProperties -MockWith {$True}

        Mock -Command Test-AllPropertiesNotNull -MockWith {$True}

        $result = Test-IsValidRegistryKeyModel $model

        $result | Should -Be $expected
    }

    It 'All expected Properties are present and not null' -ForEach @(
        #@{Model = $null ; Expected = $False}
        @{Model = @{Properties = 
            @{
                Class = "asdf"
                Assembly = "asdf"
                CodeBase = "asdf"
                RuntimeVersion = "asdf"}} ; Expected = $True},
        @{Model = @{Properties = 
            @{
                Class = $null
                Assembly = "asdf"
                CodeBase = "asdf"
                RuntimeVersion = "asdf"}} ; Expected = $False},
        @{Model = @{Properties = 
            @{
                Assembly = "asdf"
                CodeBase = "asdf"
                RuntimeVersion = "asdf"}} ; Expected = $False},
        @{Model = @{Properties = 
            @{
                Class = "asdf"
                Assembly = $null
                CodeBase = "asdf"
                RuntimeVersion = "asdf"}} ; Expected = $False},
        @{Model = @{Properties = 
            @{
                Class = "asdf"
                CodeBase = "asdf"
                RuntimeVersion = "asdf"}} ; Expected = $False},
        @{Model = @{Properties = 
            @{
                Class = "asdf"
                Assembly = "asdf"
                CodeBase = $null
                RuntimeVersion = "asdf"}} ; Expected = $False},
        @{Model = @{Properties = 
            @{
                Class = "asdf"
                Assembly = "asdf"
                RuntimeVersion = "asdf"}} ; Expected = $False}
        @{Model = @{Properties = 
            @{
                Class = "asdf"
                Assembly = "asdf"
                CodeBase = "asdf"
                RuntimeVersion = $null}} ; Expected = $False},
        @{Model = @{Properties = 
            @{
                Class = "asdf"
                Assembly = "asdf"
                CodeBase = "asdf"}} ; Expected = $False}
    ) {
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1 Test-HasAllProperties
        
        Mock -Command Test-HasAllKeys -MockWith {$True}

        Mock -Command Test-AllKeysNotNull -MockWith {$True}

        $result = Test-IsValidRegistryKeyModel $model

        $result | Should -Be $expected
    }
}
