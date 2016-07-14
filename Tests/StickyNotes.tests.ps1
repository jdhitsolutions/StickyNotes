$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. $here\..\$sut

Describe StickyNotes {
    It "does something useful" {
        $true | Should Be $true
    }
}
