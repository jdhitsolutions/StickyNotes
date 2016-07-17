#requires -version 5.0

#functions for the StickyNotes module

Function Restore-StickyNote {
[cmdletbinding()]
Param()

Write-Verbose "Starting: $($MyInvocation.Mycommand)"

[myStickyNote]::new()

Write-Verbose "Ending: $($MyInvocation.Mycommand)"

}

Function New-StickyNote {

[cmdletbinding()]
Param(
[Parameter(
Position=0,
Mandatory,
HelpMessage="Enter text for the sticky note"
)]
[string]$Text,
[switch]$Bold,
[switch]$Italic,
[switch]$Underline,
[ValidateSet("Left","Center","Right")]
[string]$Alignment = "Left"
)

Write-Verbose "Starting $($MyInvocation.MyCommand)"
#display PSBoundparameters formatted nicely for Verbose output  
[string]$pb = ($PSBoundParameters | Format-Table -AutoSize | Out-String).TrimEnd()
Write-Verbose "PSBoundparameters: `n$($pb.split("`n").Foreach({"$("`t"*4)$_"}) | Out-String) `n" 

[mystickynote]::NewNote($Text,$Alignment,0,$Bold,$underline,$Italic)

Write-Verbose "Ending $($MyInvocation.MyCommand)"


} #end function

Function Set-StickyNote {

[cmdletbinding()]
Param(
[Parameter(position=0)]
[string]$Text,
[switch]$Bold,
[switch]$Italic,
[switch]$Underline,
[ValidateSet("Left","Center","Right")]
[string]$Alignment,
[switch]$Append,
[int]$FontSize = 0
)

Write-Verbose "Starting $($MyInvocation.MyCommand)"
#display PSBoundparameters formatted nicely for Verbose output  
[string]$pb = ($PSBoundParameters | Format-Table -AutoSize | Out-String).TrimEnd()
Write-Verbose "PSBoundparameters: `n$($pb.split("`n").Foreach({"$("`t"*4)$_"}) | Out-String) `n" 

if ($Text -AND $append) {
Write-Verbose "Appending"
    [MyStickyNote]::SetText($Text,$True)
}
elseif ($Text) {
   [MyStickyNote]::SetText($Text,$False)
}

if ($Alignment) {
    [MyStickyNote]::SetAlignment($Alignment)
}

if ($Bold) {
    [MyStickyNote]::ToggleBold()
}

if ($Underline) {
    [MyStickyNote]::ToggleUnderline()
}
if ($Italic) {
    [MyStickyNote]::ToggleItalic()
}

if ($FontSize) {
    [MyStickyNote]::SetFontSize($FontSize)
}

Write-Verbose "Ending $($MyInvocation.MyCommand)"


} #end function

Function Remove-StickyNote {

[cmdletbinding(SupportsShouldProcess)]
Param()

Write-Verbose "Starting $($MyInvocation.MyCommand)"
if ([mystickynote]::TestProcess()) {
    if ($PSCmdlet.ShouldProcess("Sticky Notes")) {
    [mystickynote]::Remove()
    }
}
else {
    Write-Warning "No sticky notes seem to be running"
}

Write-Verbose "Ending $($MyInvocation.MyCommand)"


} #end function
