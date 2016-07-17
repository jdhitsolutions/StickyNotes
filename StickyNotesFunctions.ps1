#requires -version 5.0

#functions for the StickyNotes module

Function Restore-StickyNote {
[cmdletbinding()]
Param()

Write-Verbose "Starting: $($MyInvocation.Mycommand)"

[myStickyNote]::new()

Write-Verbose "Ending: $($MyInvocation.Mycommand)"

} #end function

Function New-StickyNote {

[cmdletbinding()]
Param(
[Parameter(
Position=0,
Mandatory,
HelpMessage="Enter text for the sticky note",
ValueFromPipeline
)]
[string]$Text,
[switch]$Bold,
[switch]$Italic,
[switch]$Underline,
[ValidateSet("Left","Center","Right")]
[string]$Alignment = "Left"
)

Begin {
    Write-Verbose "[BEGIN  ] Starting $($MyInvocation.MyCommand)"
    #display PSBoundparameters formatted nicely for Verbose output  
    [string]$pb = ($PSBoundParameters | Format-Table -AutoSize | Out-String).TrimEnd()
    Write-Verbose "[BEGIN  ] PSBoundparameters: `n$($pb.split("`n").Foreach({"$("`t"*4)$_"}) | Out-String) `n" 
}

Process {
    Write-Verbose "[PROCESS] Creating a sticky note."
    [mystickynote]::NewNote($Text,$Alignment,0,$Bold,$underline,$Italic)
}

End {
    Write-Verbose "[END    ] Ending $($MyInvocation.MyCommand)"
}

} #end function

Function Set-StickyNote {

[cmdletbinding()]
Param(
[Parameter(Position=0,ValueFromPipeline)]
[string]$Text,
[switch]$Bold,
[switch]$Italic,
[switch]$Underline,
[ValidateSet("Left","Center","Right")]
[string]$Alignment,
[switch]$Append,
[int]$FontSize = 0
)

Begin {
    Write-Verbose "[BEGIN  ] Starting $($MyInvocation.MyCommand)"
    #display PSBoundparameters formatted nicely for Verbose output  
    [string]$pb = ($PSBoundParameters | Format-Table -AutoSize | Out-String).TrimEnd()
    Write-Verbose "[BEGIN  ] PSBoundparameters: `n$($pb.split("`n").Foreach({"$("`t"*4)$_"}) | Out-String) `n" 
}

Process {

    if ($Text -AND $append) {
        Write-Verbose "[PROCESS] Appending text"
        [MyStickyNote]::SetText($Text,$True)
    }
    elseif ($Text) {
        Write-Verbose "[PROCESS] Replacing text"
       [MyStickyNote]::SetText($Text,$False)
    }

    if ($Alignment) {
        Write-Verbose "[PROCESS] Aligning $Alignment"
        [MyStickyNote]::SetAlignment($Alignment)
    }

    if ($Bold) {
        Write-Verbose "[PROCESS] Toggling Bold"
        [MyStickyNote]::ToggleBold()
    }

    if ($Underline) {
        Write-Verbose "[PROCESS] Toggling Underline"
        [MyStickyNote]::ToggleUnderline()
    }
    if ($Italic) {
        Write-Verbose "[PROCESS] Toggling Italic"
        [MyStickyNote]::ToggleItalic()
    }

    if ($FontSize) {
        Write-Verbose "[PROCESS] Modifying Font Size $Fontsize"
        [MyStickyNote]::SetFontSize($FontSize)
    }
 }
   
End {
    Write-Verbose "[END    ] Ending $($MyInvocation.MyCommand)"
}

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
