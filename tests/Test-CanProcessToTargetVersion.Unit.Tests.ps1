Describe 'Test-CanProcessToTargetVersion tests'{
    BeforeAll {        
        . $PSScriptRoot\..\src\Switch-RDVersionCode.ps1
        Initialize-ScriptScopeFields @('2.5.2.1', '3.0.0.0')    
    }
    Context 'Config File is Missing'{
        BeforeAll {
            Mock -Command Test-Path -MockWith {$False}
        }
        It 'Fails if target is RD3 and RD3 config data file is missing'{
        
            $values = Test-CanProcessToTargetVersion $false
    
            $values.CanProcess | Should -Be $False
        }
    
        It 'Succeeds if target is RD2 and RD3 config data file is missing'{
            
            Mock -CommandName Test-Path -ParameterFilter {$Path -like "*2.*"} -MockWith {$True}
            
            Mock -Command Test-IsConfiguredForRD3 -MockWith {$True}
            
            Mock -Command Test-IsConfiguredForRD2 -MockWith {$False}

            Mock -Command Get-RD2PrototypeKeys -MockWith {
                @("TestPathExt\2.x.x.x", "TestPathDock\2.x.x.x")
            }
    
            $values = Test-CanProcessToTargetVersion $True
    
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
    
            $values = Test-CanProcessToTargetVersion $false
    
            $values.CanProcess | Should -Be $False
        }
    
        It 'Succeeds if target is RD2 and RD3 config data file is Empty'{
            
            Mock -Command Test-IsConfiguredForRD2 -MockWith {$False}
            Mock -Command Test-IsConfiguredForRD3 -MockWith {$True}
    
            $values = Test-CanProcessToTargetVersion $True
    
            $values.CanProcess | Should -Be $True
        }
    }

    It 'Expected = <expected> if target is RD3 and config data file yields <fileReadResult.Count> models' -ForEach @(
        @{FileReadResult = @() ; Expected = $False}
        @{FileReadResult = @("Something") ; Expected = $False}
        @{FileReadResult = @("Something", "Else") ; Expected = $True}
        @{FileReadResult = @("Something", "Else", "AndMore") ; Expected = $False}
    ){
        Initialize-ScriptScopeFields @('2.5.2.0')  #starting configuration is RD2  
        
        Mock -CommandName New-RegistryKeyModelsFromFile -MockWith {$FileReadResult} 

        Mock -CommandName Test-Path -MockWith {$true}

        Mock -Command Test-FileIsLocked -MockWith {$false}

        $values = Test-CanProcessToTargetVersion $false

        $values.CanProcess | Should -Be $expected
    }
    It 'RD3 config data file IsLocked = <islocked> and IsRD2Target = <isRd2Target>'  -ForEach @(
        #read-only is OK for setting up RD3
        @{IsLocked = $True ; IsRD2Target = $False ; Expected = $False} #Changing RD2 to RD3
        #file locked just means it will have old data...but go ahead and setup RD2
        @{IsLocked = $True ; IsRD2Target = $True ; Expected = $False} #Changing RD2 to RD2 (no-op message)
        @{IsLocked = $False ; IsRD2Target = $False ; Expected = $True} #Changing RD2 to RD3
        @{IsLocked = $False ; IsRD2Target = $True ; Expected = $False} #Changing RD2 to RD2 (no-op message)
    ){
        
        Mock -Command Test-Path -MockWith {$True}

        Mock -CommandName Test-FileIsLocked -MockWith {
            $isLocked
        } 

        Mock -CommandName New-RegistryKeyModelsFromFile -MockWith { @("KeyOne", "KeyTwo") }

        $values = Test-CanProcessToTargetVersion $IsRD2Target

        $values.CanProcess | Should -Be $expected
    }

    It 'RD2 prototype key Exists = <exists> IsRD2Target = <isRd2Target>' -ForEach @(
        @{Exists = $True ; IsRD2Target = $False ; Expected = $False} #Loading RD3 when RD3 is active is a no-op
        @{Exists = $True ; IsRD2Target = $True ; Expected = $True}
        #Not going to happen...RD2 is not installed - but it's absence wouldn't prevent installing RD3
        @{Exists = $False ; IsRD2Target = $False ; Expected = $False} #Loading RD3 when RD3 is active is a no-op 
        @{Exists = $False ; IsRD2Target = $True ; Expected = $False}
    ){
        
        Mock -CommandName Test-Path -ParameterFilter {-not ($Path -like "*2.*") } -MockWith {
            $True
        }  

        Mock -CommandName Test-Path -ParameterFilter {$Path -like "*2.*" } -MockWith {
            $Exists
        } 
        
        Mock -Command Get-RD2PrototypeKeys -MockWith {@("TestPathExt\2.x.x.x", "TestPathDock\2.x.x.x")}

        if ($Exists){
            Initialize-ScriptScopeFields @('2.5.2.0', '3.0.0.0')   
        } else {
            Initialize-ScriptScopeFields @('3.0.0.0')   
        }

        $values = Test-CanProcessToTargetVersion $IsRD2Target

        $values.CanProcess | Should -Be $expected
    }

    It 'Fails if target is RD3 and RD3 config data is not JSON'{
        
        Mock -Command Test-Path -MockWith {$True}

        Mock -CommandName New-RegistryKeyModelsFromFile -MockWith {$null} 

        $values = Test-CanProcessToTargetVersion $false

        $values.CanProcess | Should -Be $False
    }
}
