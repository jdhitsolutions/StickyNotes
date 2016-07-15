#requires -version 5.0

#StickyNotes.psm1

Enum Alignment {
    Left
    Right
    Center
}

Class MyStickyNote {

# this class has no properties

#methods

static [void]Remove() {
    #$title = $this.ResolveProcess($this.ID) 
    [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
    [System.Windows.Forms.SendKeys]::SendWait("^d")
    [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
}

static [void]ToggleBold() {
  #$title = $this.ResolveProcess($this.ID) 
  [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
  [System.Windows.Forms.SendKeys]::SendWait("^a")
  Start-Sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("^b")
  Start-Sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("{Down}")
  [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
}

static [void]ToggleItalic() {
  
  [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
  [System.Windows.Forms.SendKeys]::SendWait("^a")
  Start-Sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("^i")
  Start-Sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("{Down}")
  [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
}

static [void]ToggleUnderline() {
   
  [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
  [System.Windows.Forms.SendKeys]::SendWait("^a")
  start-sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("^u")
  start-sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("{Down}")
  [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
  
}

static [boolean]TestProcess() {
  if (Get-Process -Name StikyNot -ErrorAction SilentlyContinue) {
    Return $True
  }
  else {
    Return $False
  }
}

static [void]SetAlignment([Alignment]$Alignment) {
  
    [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
     Start-Sleep -Milliseconds 50
    [System.Windows.Forms.SendKeys]::SendWait("^a")
   
    Switch ($Alignment) {
        "Left" {
            [System.Windows.Forms.SendKeys]::SendWait("^l")
        }

        "Center" {
            [System.Windows.Forms.SendKeys]::SendWait("^e")
        }
        "Right" {
            [System.Windows.Forms.SendKeys]::SendWait("^r")
        }
    } #switch

    [System.Windows.Forms.SendKeys]::SendWait("{Down}")
    [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
}

static [void]SetFontSize([int]$Size) {
     
     [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
     [System.Windows.Forms.SendKeys]::SendWait("^a")
     Start-Sleep -Milliseconds 50

    if ($Size -lt 0) {
        0..($Size -1) | foreach {
        [System.Windows.Forms.SendKeys]::SendWait("^+<")
        }
    }
    elseif ($Size -gt 0) {
        0..($Size-1) | foreach {
        [System.Windows.Forms.SendKeys]::SendWait("^+>")
        }
    }
    
    Start-Sleep -Milliseconds 50
    [System.Windows.Forms.SendKeys]::SendWait("{Down}")
    [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
}

static [void]SetText([string]$Text,[boolean]$Append) {

    $text | Set-Clipboard
    
    [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
    if (-Not $Append) {
        [System.Windows.Forms.SendKeys]::SendWait("^a")
    }

    [System.Windows.Forms.SendKeys]::SendWait("^v")
    #give the paste a moment to complete
    Start-Sleep -Milliseconds 50
    [System.Windows.Forms.SendKeys]::SendWait("{Enter}")

    #[System.Windows.Forms.SendKeys]::SendWait("{Down}")
    [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
    
}

static [void]NewNote ([string]$Text,[alignment]$Alignment="Left",[int]$FontSize=0,[boolean]$Bold,[boolean]$Underline,[boolean]$Italic) {

if (-Not ([myStickyNote]::TestProcess())) {
    Start-Process -FilePath Stikynot.exe
    Start-Sleep -Milliseconds 150
}
else {
  [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
  Start-Sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("^n")
}

  #copy text to clipboard and paste it. Faster than trying to send keys
  $text | Set-Clipboard
  
  [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
  [System.Windows.Forms.SendKeys]::SendWait("^v")
  [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
  #give the paste a moment to complete
  Start-Sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("{Enter}")

  [MyStickyNote]::SetAlignment($Alignment)
  [MyStickyNote]::SetFontSize($FontSize)
  if ($bold) { [MyStickyNote]::ToggleBold()}
  if ($Underline) { [MyStickyNote]::ToggleUnderline()}
  if ($Italic) { [MyStickyNote]::ToggleItalic()}
  
  [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
}

#constructor

MyStickyNote() {

if (-Not ([myStickyNote]::TestProcess())) {
    Start-Process -FilePath Stikynot.exe
    Start-Sleep -Milliseconds 150
}
else {
    [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
    [System.Windows.Forms.SendKeys]::SendWait("^n")
}
[Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)


} #New()


} #class

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

#define aliases
Set-Alias -Name kn -Value Remove-StickyNote
Set-Alias -Name sn -Value Set-StickyNote
Set-Alias -Name nn -Value New-StickyNote
Set-Alias -Name rn -Value Restore-StickyNote


