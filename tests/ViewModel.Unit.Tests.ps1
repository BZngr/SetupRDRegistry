Describe 'Unresolved Version(s)'{
    BeforeAll{
        . $PSScriptRoot\..\src\Switch-RDVersionCode.ps1 New-ViewModel
    }   
    
    It 'Generates correct ErrorSource' -ForEach @(
        @{Act = "2.5.4.3" ; InAct = "3.0.0.0" ; RD2 = "2.5.4.3" ; RD3 = "3.0.0.0" ; Expected = ""}
        @{Act = "3.0.0.0" ; InAct = "2.5.4.3" ; RD2 = "2.5.4.3" ; RD3 = "3.0.0.0" ; Expected = ""}
        @{Act = "2.5.4.3" ; InAct = "TBD" ; RD2 = "2.5.4.3" ; RD3 = "TBD" ; Expected = "3"}
        @{Act = "3.0.0.0" ; InAct = "TBD" ; RD2 = "TBD" ; RD3 = "3.0.0.0" ; Expected = "2"}
        @{Act = "TBD" ; InAct = "TBD" ; RD2 = "TBD" ; RD3 = "TBD" ; Expected = "NO RD VERSION REGISTERED"}
    ){
        $sut = New-ViewModel $act $inact $rd2 $rd3

        $sut.CV | Should -Be $act
        $sut.NV | Should -Be $inact
        $sut.ErrorVersion | Should -Be $expected
    }

    It 'Generates correct Primary Message' -ForEach @(
        @{Act = "2.5.4.3" ; InAct = "3.0.0.0" ; RD2 = "2.5.4.3" ; RD3 = "3.0.0.0" ; Expected = ""}
        @{Act = "3.0.0.0" ; InAct = "2.5.4.3" ; RD2 = "2.5.4.3" ; RD3 = "3.0.0.0" ; Expected = ""}
        @{Act = "2.5.4.3" ; InAct = "TBD" ; RD2 = "2.5.4.3" ; RD3 = "TBD" ; Expected = "Unable to locate v3 setup data"}
        @{Act = "3.0.0.0" ; InAct = "TBD" ; RD2 = "TBD" ; RD3 = "3.0.0.0" ; Expected = "Rubberduck v2 Registry Entries not found"}
        @{Act = "TBD" ; InAct = "TBD" ; RD2 = "TBD" ; RD3 = "TBD" ; Expected = "Rubberduck version(s) data not found"}
    ){
        $sut = New-ViewModel $act $inact $rd2 $rd3

        $sut.ErrMsg.Primary | Should -Be $expected
    }

    It 'Generates correct FIx Message' -ForEach @(
        @{Act = "2.5.4.3" ; InAct = "3.0.0.0" ; RD2 = "2.5.4.3" ; RD3 = "3.0.0.0" ; Expected = ""}
        @{Act = "3.0.0.0" ; InAct = "2.5.4.3" ; RD2 = "2.5.4.3" ; RD3 = "3.0.0.0" ; Expected = ""}
        @{Act = "2.5.4.3" ; InAct = "TBD" ; RD2 = "2.5.4.3" ; RD3 = "TBD" ; Expected = "To switch to v3, please Clean\Build or Install Rubberduck v3"}
        @{Act = "3.0.0.0" ; InAct = "TBD" ; RD2 = "TBD" ; RD3 = "3.0.0.0" ; Expected = "Please Build or Install Rubberduck v2"}
        @{Act = "TBD" ; InAct = "TBD" ; RD2 = "TBD" ; RD3 = "TBD" ; Expected = "Please Install or Build the Rubberduck version of interest"}
    ){
        $sut = New-ViewModel $act $inact $rd2 $rd3

        $sut.ErrMsg.Fix | Should -Be $expected
    }
}
