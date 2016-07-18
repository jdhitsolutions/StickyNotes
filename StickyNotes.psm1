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

Enum NoteColor {
    Blue
    Green
    Pink
    Purple
    White
    Yellow
}

Class MyStickyNote {

    #this class has no properties

    #methods

    static [void]Remove() {
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        [System.Windows.Forms.SendKeys]::SendWait("^d")
        [System.Windows.Forms.SendKeys]::SendWait("%Y")
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
    }

    static [void]ToggleFormat($FormatType) {
        $Clipboard = Get-Clipboard
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        Start-Sleep -Milliseconds 150
        $RtfNote = [Windows.Clipboard]::GetData([System.Windows.Forms.DataFormats]::Rtf)
        if ($RtfNote -match "\\$FormatType")
        {
            $RtfNote = $RtfNote -replace "\\$FormatType(\d|none)?",''
        }
        else
        {
            $RtfNote = $RtfNote -replace '\\f0',"\$FormatType\f0"
        }
        [Windows.Clipboard]::SetData([System.Windows.Forms.DataFormats]::Rtf, $RtfNote)
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        Start-Sleep -Milliseconds 150
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
        $Clipboard | Set-Clipboard
    }

    static [void]ToggleBold() {
        [MyStickyNote]::ToggleFormat('b')
    }

    static [void]ToggleItalic() {  
        [MyStickyNote]::ToggleFormat('i')
    }

    static [void]ToggleUnderline() {   
        [MyStickyNote]::ToggleFormat('ul')
    }

    static [void]SetColor([NoteColor]$Color) {
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait('+{F10}')

        Switch ($Color) {
            "Blue" {
                [System.Windows.Forms.SendKeys]::SendWait("b")
                break
            }
            "Green" {
                [System.Windows.Forms.SendKeys]::SendWait("g")
                break
            }
            "Pink" {
                [System.Windows.Forms.SendKeys]::SendWait("p")
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                break
            }
            'Purple' {
                [System.Windows.Forms.SendKeys]::SendWait("p")
                [System.Windows.Forms.SendKeys]::SendWait("p")
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                break
            }
            'White' {
                [System.Windows.Forms.SendKeys]::SendWait("w")
                break
            }
            'Yellow' {
                [System.Windows.Forms.SendKeys]::SendWait("y")
                break
            }
            } #switch

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
        $Clipboard = Get-Clipboard
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        Start-Sleep -Milliseconds 150
        $RtfNote = [Windows.Clipboard]::GetData([System.Windows.Forms.DataFormats]::Rtf)
        Switch ($Alignment) {
            "Left" {
                $RtfNote = $RtfNote -replace "\\pard\\(qc|qr)",''
            }

            "Center" {
                $RtfNote = $RtfNote -replace "\\pard(\\qr)?",'\pard\qc'
            }
            "Right" {
                $RtfNote = $RtfNote -replace "\\pard(\\qc)?",'\pard\qr'
            }
        } #switch
        [Windows.Clipboard]::SetData([System.Windows.Forms.DataFormats]::Rtf, $RtfNote)
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        Start-Sleep -Milliseconds 150
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
        $Clipboard | Set-Clipboard
    }

    static [void]SetFontFamily([string]$FontFamily)
    {
        $Clipboard = Get-Clipboard
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        Start-Sleep -Milliseconds 50

        [System.Windows.Forms.SendKeys]::SendWait("^a")
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        $RtfNote = [Windows.Clipboard]::GetData([System.Windows.Forms.DataFormats]::Rtf)
        $RtfNote = $RtfNote -replace "fcharset0 [\w\s]*;}","fcharset0 $FontFamily;}" -replace "fnil [\w\s]*;}","fnil $FontFamily;}"
        [Windows.Clipboard]::SetData([System.Windows.Forms.DataFormats]::Rtf, $RtfNote)
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)

        $Clipboard | Set-Clipboard
    }

    static [string]FindNote([string]$Query,[int]$Timeout=10)
    {
        $Clipboard = Get-Clipboard
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        Start-Sleep -Milliseconds 50

        $Notes = @()
        $Timer = [System.Diagnostics.Stopwatch]::StartNew()
        $Result = ''

        while ($true) {
            [System.Windows.Forms.SendKeys]::SendWait("^a")
            [System.Windows.Forms.SendKeys]::SendWait("^c")
            [System.Windows.Forms.SendKeys]::SendWait("{Down}")
            $NoteText = Get-Clipboard

            if ($NoteText -match $Query)
            {
                $Result = $NoteText
                break
            }
            elseif ($Notes -contains $NoteText -or $Timer.Elapsed.Seconds -ge $Timeout)
            {
                break
            }
            $Notes += $NoteText

            [System.Windows.Forms.SendKeys]::SendWait("^{TAB}")
        }

        $Clipboard | Set-Clipboard

        return $Result
    }

    static [void]SetFontSize([int]$Size) {
        $Clipboard = Get-Clipboard
        # double input to match final pt size
        $Size = $Size * 2
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        Start-Sleep -Milliseconds 150
        $RtfNote = [Windows.Clipboard]::GetData([System.Windows.Forms.DataFormats]::Rtf)
        $RtfNote = $RtfNote -replace "\\fs\d\d?\d?","\fs$Size"

        [Windows.Clipboard]::SetData([System.Windows.Forms.DataFormats]::Rtf, $RtfNote)
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        Start-Sleep -Milliseconds 150
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
        $Clipboard | Set-Clipboard
    }

    static [void]SetText([string]$Text,[boolean]$Append) {    
        [Microsoft.VisualBasic.Interaction]::AppActivate("Sticky Notes")
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        $RtfNote = [Windows.Clipboard]::GetData([System.Windows.Forms.DataFormats]::Rtf)
        $TextNote = [Windows.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::Text)
        if ($Append) {
            $RtfNote = $RtfNote -replace $TextNote,"$TextNote $Text"
        }
        else
        {
            $RtfNote = $RtfNote -replace $TextNote,$Text
        }

        $RtfNote = [Windows.Clipboard]::SetData([System.Windows.Forms.DataFormats]::Rtf, $RtfNote)
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)

    }

    static [void]NewNote ([string]$Text,[alignment]$Alignment="Left",[int]$FontSize=0,[boolean]$Bold,[boolean]$Underline,[boolean]$Italic,[NoteColor]$Color="Yellow",[string]$FontFamily) {

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
        if ($FontFamily) { [MyStickyNote]::SetFontFamily($FontFamily) }
        [MyStickyNote]::SetColor($Color)

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