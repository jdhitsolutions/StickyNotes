#requires -version 5.0

#functions for the StickyNotes module

Function Restore-StickyNote {
    [CmdletBinding()]
    Param()

    Write-Verbose "Starting: $($MyInvocation.Mycommand)"

    [myStickyNote]::new()

    Write-Verbose "Ending: $($MyInvocation.Mycommand)"

} #end function

Function New-StickyNote {

    [CmdletBinding()]
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
        [string]$Alignment = "Left",
        [ValidateSet("Blue","Green","Pink","Purple","White","Yellow")]
        [string]$Color = "Yellow",
        [ValidateScript({
            (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name -contains $_
            })]
        [string]$FontFamily = "Segoe Print"
        )

    Begin {
        Write-Verbose "[BEGIN  ] Starting $($MyInvocation.MyCommand)"
        #display PSBoundparameters formatted nicely for Verbose output  
        [string]$pb = ($PSBoundParameters | Format-Table -AutoSize | Out-String).TrimEnd()
        Write-Verbose "[BEGIN  ] PSBoundparameters: `n$($pb.split("`n").Foreach({"$("`t"*4)$_"}) | Out-String) `n" 
    }

    Process {
        Write-Verbose "[PROCESS] Creating a sticky note."
        [mystickynote]::NewNote($Text,$Alignment,0,$Bold,$underline,$Italic,$Color,$FontFamily)
    }

    End {
        Write-Verbose "[END    ] Ending $($MyInvocation.MyCommand)"
    }

} #end function

Function Set-StickyNote {

    [CmdletBinding()]
    Param(
        [Parameter(Position=0,ValueFromPipeline)]
        [string]$Text,
        [switch]$Bold,
        [switch]$Italic,
        [switch]$Underline,
        [ValidateSet("Left","Center","Right")]
        [string]$Alignment,
        [ValidateSet("Blue","Green","Pink","Purple","White","Yellow")]
        [string]$Color,
        [switch]$Append,
        [ValidateScript({
            (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name -contains $_
            })]
        [string]$FontFamily,
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

        if ($Color) {
            Write-Verbose "[PROCESS] Setting note color to $Color"
            [MyStickyNote]::SetColor($Color)
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

        if ($FontFamily)
        {
            Write-Verbose "[PROCESS] Modifying Font Family $FontFamily"
            [MyStickyNote]::SetFontFamily($FontFamily)
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

Function Find-StickyNote {
    [CmdletBinding()]
    Param(
        [string]$Query,
        [int]$Timeout=10,
        [switch]$PassThru
        )
    $NoteText = [MyStickyNote]::FindNote($Query, $Timeout)
    if ($PassThru)
    {
        $NoteText
    }
}

Function Remove-StickyNote{

    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
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