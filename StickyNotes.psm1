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
     
     if (($Size -lt -5) -OR ($Size -gt 5)) { 
      Write-Warning "You just specify a size between -5 and 5"
      #bail out
      Return
      
      }

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
    <#
    else {
        #restore to original size
        if ($this.FontSize -lt 0) {
            Do {
                [System.Windows.Forms.SendKeys]::SendWait("^+>")
                $this.FontSize++
            } until ($this.FontSize -eq 0)
        }
        elseif ($this.FontSize -gt 0) {
            Do {
                [System.Windows.Forms.SendKeys]::SendWait("^+<")
                $this.FontSize--
            } until ($this.FontSize -eq 0)
        }
    }
    #>
    Start-Sleep -Milliseconds 50
    [System.Windows.Forms.SendKeys]::SendWait("{Down}")
    [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
    #$this.FontSize += $size
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

Function New-StickyNote {

<#
.Synopsis
Create a new sticky note.
.Description
This command will create a sticky note. You can specify additional formatting options. 

.Example
PS C:\> New-StickyNote "pickup milk on the way home" -bold
.Notes
Version     : 1.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

#>

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

<#
.Synopsis
Set text and style for a sticky note.
.Description
This command will set the text and style for a sticky note. If there are multiple notes, this command will set the one that has focus, which you must do manually. The default is the last note created.

The style parameters like Bold behave more like toggles. If text is already in bold than using -Bold will turn it off and vice versa.

.Parameter Append
Append your text to the end of the note. Otherwise existing text will be replaced
.Example
PS C:\> set-stickynote "pickup milk on the way home" -bold -underline -append
.Notes

Version     : 1.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/


#>

[cmdletbinding()]
Param(
[Parameter(position=0)]
[string]$Text,
[switch]$Bold,
[switch]$Italic,
[switch]$Underline,
[ValidateSet("Left","Center","Right")]
[string]$Alignment,
[switch]$Right,
[switch]$Append,
[ValidateRange(-5,5)]
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

<#
.Synopsis
Remove a new sticky note.
.Description
This command will remove the sticky note that has focus. If you have more than one note, the last one created has focus, unless you manually select a different one.

If you kill the stikynot process, the next time you create a note, any previously created notes will return.

.Example
PS C:\> remove-stickynote 

.Notes
Version     : 1.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/


#>

[cmdletbinding(SupportsShouldProcess)]
Param()

Write-Verbose "Starting $($MyInvocation.MyCommand)"
if ([mystickynote]::TestProcess()) {
    [mystickynote]::Remove()
}
else {
    Write-Warning "No sticky notes seem to be running"
}

Write-Verbose "Ending $($MyInvocation.MyCommand)"


} #end function

#define aliases
Set-Alias -Name rn -Value Remove-StickyNote
Set-Alias -Name sn -Value Set-StickyNote
Set-Alias -Name nn -Value New-StickyNote


