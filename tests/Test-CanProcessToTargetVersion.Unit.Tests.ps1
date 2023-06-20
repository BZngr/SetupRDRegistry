Describe 'Test-CanProcessToTargetVersion tests'{
    BeforeAll {        
        . $PSScriptRoot\..\src\SetRegistryForRDVersionImpl.ps1

        function Get-TestKeys($isRD2Target){
            if ($IsRD2Target){
                @("RegKey1", "RegKey2")
            } else {@()}
        }
    }
    Context 'Config File is Missing'{
        BeforeAll {
            Mock -Command Test-Path -MockWith {$False}
        }
        It 'Fails if target is RD3 and RD3 config data file is missing'{
        
            $values = Test-CanProcessToTargetVersion (Get-TestKeys($False)) $false
    
            $values.CanProcess | Should -Be $False
        }
    
        It 'Succeeds if target is RD2 and RD3 config data file is missing'{
            
            Mock -CommandName Test-Path -ParameterFilter {$Path -like "*2.*"} -MockWith {$True}
            
            Mock -Command Test-IsConfiguredForRD3 -MockWith {$True}
            
            Mock -Command Test-IsConfiguredForRD2 -MockWith {$False}

            Mock -Command Get-RD2PrototypeKeys -MockWith {@("TestPathExt\2.x.x.x", "TestPathDock\2.x.x.x")}
    
            $values = Test-CanProcessToTargetVersion (Get-TestKeys($True)) $True
    
            $values.CanProcess | Should -Be $True
        }
    }
    Context 'Config File is present' {
        BeforeAll {
            Mock -Command Test-Path -MockWith {$True}
        }
        It 'Fails if target is RD3 and RD3 config data file is Empty'{
        
            Mock -Command Get-CachedRD3KeyValuesFilepath -MockWith {".\tests\TempEmptyRD3Values.txt"}
    
            Mock -Command Test-IsConfiguredForRD2 -MockWith {$True}
            Mock -Command Test-IsConfiguredForRD3 -MockWith {$False}
    
            $values = Test-CanProcessToTargetVersion @() $false
    
            $values.CanProcess | Should -Be $False
        }
    
        It 'Succeeds if target is RD2 and RD3 config data file is Empty'{
            
            Mock -Command Test-IsConfiguredForRD2 -MockWith {$False}
            Mock -Command Test-IsConfiguredForRD3 -MockWith {$True}
    
            $values = Test-CanProcessToTargetVersion @("RegKey1", "RegKey2") $True
    
            $values.CanProcess | Should -Be $True
        }
    }

    It 'Expected = <expected> if target is RD3 and config data file yields <fileReadResult.Count> models' -ForEach @(
        @{FileReadResult = @() ; Expected = $False}
        @{FileReadResult = @("Something") ; Expected = $False}
        @{FileReadResult = @("Something", "Else") ; Expected = $True}
        @{FileReadResult = @("Something", "Else", "AndMore") ; Expected = $False}
    ){
        Mock -CommandName Test-IsConfiguredForRD2 -MockWith {$True}
        Mock -CommandName Test-IsConfiguredForRD3 -MockWith {$False}
        
        Mock -CommandName New-MaybeRegistryKeyModelsFromFile -MockWith {
            $FileReadResult
        } 

        $values = Test-CanProcessToTargetVersion (Get-TestKeys $IsRD2Target) $false

        $values.CanProcess | Should -Be $expected
    }

    It 'RD3 config data file IsLocked = <islocked> and IsRD2Target = <isRd2Target>' -ForEach @(
        #read-only is OK for setting up RD3
        @{IsLocked = $True ; IsRD2Target = $False ; Expected = $True} 
        #file locked just means it will have old data...but go ahead and setup RD2
        @{IsLocked = $True ; IsRD2Target = $True ; Expected = $True}
        @{IsLocked = $False ; IsRD2Target = $False ; Expected = $True}
        @{IsLocked = $False ; IsRD2Target = $True ; Expected = $True}
    ){
        
        Mock -Command Test-Path -MockWith {$True}

        Mock -CommandName Test-FileIsLocked -MockWith {
            $isLocked
        } 

        $values = Test-CanProcessToTargetVersion (Get-TestKeys $IsRD2Target) $IsRD2Target

        $values.CanProcess | Should -Be $expected
    }

    It 'RD2 prototype key Exists = <exists> IsRD2Target = <isRd2Target>' -ForEach @(
        @{Exists = $True ; IsRD2Target = $False ; Expected = $True}
        @{Exists = $True ; IsRD2Target = $True ; Expected = $True}
        #Not going to happen...RD2 is not installed - but it's absence wouldn't prevent installing RD3
        @{Exists = $False ; IsRD2Target = $False ; Expected = $True} 
        @{Exists = $False ; IsRD2Target = $True ; Expected = $False}
    ){
        
        Mock -CommandName Test-Path -ParameterFilter {-not ($Path -like "*2.*") } -MockWith {
            #Write-Verbose "Called Mock for non '*2.*' path -> Test-Path" -Verbose
            $True
        }  

        Mock -CommandName Test-Path -ParameterFilter {$Path -like "*2.*" } -MockWith {
            #Write-Verbose "Called Mock for '*2.*' path -> Test-Path" -Verbose
            $Exists
        } 
        
        Mock -Command Get-RD2PrototypeKeys -MockWith {@("TestPathExt\2.x.x.x", "TestPathDock\2.x.x.x")}

        $values = Test-CanProcessToTargetVersion (Get-TestKeys $IsRD2Target) $IsRD2Target

        $values.CanProcess | Should -Be $expected
    }

    It 'Fails if target is RD3 and RD3 config data is not JSON'{
        
        Mock -Command Test-Path -MockWith {$True}

        Mock -CommandName New-MaybeRegistryKeyModelsFromFile -MockWith {$null} 

        $values = Test-CanProcessToTargetVersion (Get-TestKeys($False)) $false

        $values.CanProcess | Should -Be $False
    }
}
