#requires -version 5.0

#StickyNotes.psm1

#verify Stikynotes.exe and throw an error if not found
If (-Not (Test-path $env:windir\system32\stikynot.exe)) {
    Throw "Cannot find $env:windir\system32\stikynot.exe. This module is intended for client operating systems." 
}

#an enumeration that can be used in the functions
Enum Alignment {
    Left
    Right
    Center
}

Class MyStickyNote {

#this class has no properties

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
  Start-Sleep -Milliseconds 50
  [System.Windows.Forms.SendKeys]::SendWait("^u")
  Start-Sleep -Milliseconds 50
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
    #use the clipboard to insert text as it is faster
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

#test if program is running and start it if it isn't.
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

  #set formatting by invoking class methods
  [MyStickyNote]::SetAlignment($Alignment)
  [MyStickyNote]::SetFontSize($FontSize)
  if ($bold) { [MyStickyNote]::ToggleBold()}
  if ($Underline) { [MyStickyNote]::ToggleUnderline()}
  if ($Italic) { [MyStickyNote]::ToggleItalic()}
  
  #jump back to the PowerShell console
  [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
}

#constructor

#This will create a blank sticky note
MyStickyNote() {

    if (-Not ([myStickyNote]::TestProcess())) {
        Start-Process -FilePath Stikynot.exe
        Start-Sleep -Milliseconds 150
    }
    else {
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        [System.Windows.Forms.SendKeys]::SendWait("^n")
    }
    #jump back to the PowerShell console
    [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)

} #New


} #class

#dot source external functions
. $PSScriptRoot\StickyNotesFunctions.ps1

#define some aliases
Set-Alias -Name kn -Value Remove-StickyNote
Set-Alias -Name sn -Value Set-StickyNote
Set-Alias -Name nn -Value New-StickyNote
Set-Alias -Name rn -Value Restore-StickyNote


