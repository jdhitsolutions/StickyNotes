#requires -version 5.0

#StickyNotes.psm1

#verify note running a Server OS and throw an error if not found
$OS = (Get-CimInstance -class Win32_Operatingsystem -Property Caption).Caption
If ( $OS -match "Server|Windows 10" ) {
    Throw "This module is intended for pre-Windows 10 client operating systems." 
}

#define a variable to reference the application 
    $global:appName = "Sticky Notes"
    $global:appProc = "stikynot"
    $global:AppLaunch = [scriptblock]::Create("$env:windir\system32\stikynot.exe")

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
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
        [System.Windows.Forms.SendKeys]::SendWait("^d")
        [System.Windows.Forms.SendKeys]::SendWait("%Y")
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
    }

    static [void]ToggleFormat($FormatType) {
        
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
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
       # [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        Start-Sleep -Milliseconds 150
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
        
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


    static [boolean]TestProcess() {
        $proc = (Get-Process -name $global:appproc -ErrorAction SilentlyContinue)
          if ($proc) {
            Return $True
        }
        else {
            Return $False
        }
    }

    static [void]SetAlignment([Alignment]$Alignment) {
     
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
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
        #[System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        Start-Sleep -Milliseconds 150
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
    }

    static [void]SetFontFamily([string]$FontFamily)
    {
      
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
        Start-Sleep -Milliseconds 50

        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        $Clipboard = Get-Clipboard
        Start-Sleep -Milliseconds 50
        $RtfNote = [Windows.Clipboard]::GetData([System.Windows.Forms.DataFormats]::Rtf)
         Start-Sleep -Milliseconds 50
        $RtfNote = $RtfNote -replace "fcharset0 [\w\s]*;}","fcharset0 $FontFamily;}" -replace "fnil [\w\s]*;}","fnil $FontFamily;}"
        Start-Sleep -Milliseconds 50
        if ($rtfNote) {
        [Windows.Clipboard]::SetData([System.Windows.Forms.DataFormats]::Rtf, $RtfNote)
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        Start-Sleep -Milliseconds 50
        #[System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        }
        else {
            Write-Warning "There was a problem changing the font to $FontFamily"
        }
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)

    }

    static [string]FindNote([string]$Query,[int]$Timeout=10)
    {
        
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
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

        return $Result
    }

    static [void]SetFontSize([int]$Size) {
        
        # double input to match final pt size
        $Size = $Size * 2
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^a")
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        $Clipboard = Get-Clipboard
        Start-Sleep -Milliseconds 150
        $RtfNote = [Windows.Clipboard]::GetData([System.Windows.Forms.DataFormats]::Rtf)
        $RtfNote = $RtfNote -replace "\\fs\d\d?\d?","\fs$Size"
        if ($RtfNote) {
        [Windows.Clipboard]::SetData([System.Windows.Forms.DataFormats]::Rtf, $RtfNote)
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        #[System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
        Start-Sleep -Milliseconds 150
        }
        else {
            Write-Warning "There was a problem setting font size"
        }
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
      }

    static [void]SetText([string]$Text,[boolean]$Append) {    

        $Text | Set-Clipboard
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("^a")
         Start-Sleep -Milliseconds 50

        if ($Append) {      
            [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
        }
        
        #Paste the new text        
        [System.Windows.Forms.SendKeys]::SendWait("^v")     
        Start-Sleep -Milliseconds 50
        [System.Windows.Forms.SendKeys]::SendWait("{Enter}")
        Start-Sleep -Milliseconds 50
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)

    }

    static [void]NewNote ([string]$Text,[alignment]$Alignment="Left",[int]$FontSize=0,[boolean]$Bold,[boolean]$Underline,[boolean]$Italic,[string]$FontFamily) {

        #test if program is running and start it if it isn't.
        if (-Not ([myStickyNote]::TestProcess())) {       
            Invoke-Command -ScriptBlock $global:AppLaunch
            do {
             Start-Sleep -Milliseconds 50
            } Until ( [myStickyNote]::TestProcess())
        }
        else {           
            [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
            Start-Sleep -Milliseconds 50
            [System.Windows.Forms.SendKeys]::SendWait("^n")
        }

        #copy text to clipboard and paste it. Faster than trying to send keys
        $text | Set-Clipboard

        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
        #give the paste a moment to complete
        Start-Sleep -Milliseconds 50
        
        #set formatting by invoking class methods
        [MyStickyNote]::SetAlignment($Alignment)
        Start-Sleep -Milliseconds 50
        #only adjust fontsize if specified and not 0
        if ($FontSize -ne 0 -AND $fontSize) {
            [MyStickyNote]::SetFontSize($FontSize)
        }
        if ($bold) { [MyStickyNote]::ToggleBold()}
        if ($Underline) { [MyStickyNote]::ToggleUnderline()}
        if ($Italic) { [MyStickyNote]::ToggleItalic()}
        if ($FontFamily) { [MyStickyNote]::SetFontFamily($FontFamily) }
            
        [System.Windows.Forms.SendKeys]::SendWait("{Enter}")
        Start-Sleep -Milliseconds 50
        #jump back to the PowerShell console
        [Microsoft.VisualBasic.Interaction]::AppActivate($global:pid)
    }

    #constructor

    #This will create a blank sticky note
    MyStickyNote() {

        if (-Not ([myStickyNote]::TestProcess())) {
            Invoke-Command -ScriptBlock $global:AppLaunch
            do {
             Start-Sleep -Milliseconds 50
            } Until ( [myStickyNote]::TestProcess())
        }
        else {
            [Microsoft.VisualBasic.Interaction]::AppActivate($global:appName)
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