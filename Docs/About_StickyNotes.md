# StickyNotes
## about_StickyNotes
            

# SHORT DESCRIPTION
The StickyNotes module is a set of PowerShell functions designed to make it
easier to work with the Stikynote.exe command which will create sticky
notes on your client desktop. This utility was introduced in Windows 7 and
should be available on all client versions of Windows since. Although 
testing with Windows 10 has been problematic and is not support for now.
    
Sadly, the legacy *Stikynot.exe* application does not have any sort of API 
or managed interface which makes it difficult to manage programmatically. 
The workarounds in this module are admittedly a "hack".

The functions in this module are designed to make it easier to create,
modify and remove sticky notes from a PowerShell prompt. 

# LONG DESCRIPTION
The commands in this module are based on a PowerShell-defined class although
you can not access the class directly. The class contains no properties and
all of the methods are static. The module commands make it easier to invoke
these methods. 

## CREATING A STICKY NOTE
Use the New-StickyNote command to create a new note. If the StickyNotes
program is not running it will be started and a new note defined. If the
program is already running, this command will create a new note. 

New-StickyNote includes parameters to adjust the formatting. You can elect
to make the font bold, italic or underline, including any combination. You
can adjust the alignment to be left, center or right. The default is Left.
```
PS C:\> New-StickyNote "Reboot CHI-DC04 @7PM" -Bold -Italic -Alignment Center
```

## MODIFYING A STICKY NOTE
You can modify the top sticky note with Set-StickyNote. There is no command
for "getting" a sticky note. You can set the text of the note and optionally
append it. You can toggle style settings like Bold. If the note is already
in Bold using the -Bold parameter will disable bold and vice versa.

You can also increase or descrease the font size with a positive or negative
integer. The integer is the number of times the font size will be increased
or decreased.
```
PS C:\> Set-StickyNote -Bold -FontSize 2
```

## REMOVING A STICKY NOTE
To delete a sticky note use the Remove-StickyNote command. This will remove
the top most sticky note, if more than one. This means that if you have more
than one note you'll need to manually select the one you want to remove.
```
PS C:\> Remove-StickyNote

```
# NOTE
Be aware that all notes run under the same process and there is no way to 
access a specific note. Any actions are performed with the note which has
focus. This is typically the top most note. If you have multiple notes you
might need to manually select the correct one before running any of the 
commands in this module.

There is no way to change the size, position or color of the sticky note,
although you can manually make these changes. If you change the note color
new notes in the same process will use this new color.

If you kill the stikynot.exe process, any existing notes will be restored 
the next time you create a new note. If you think you might have existing
notes use the Restore-StickyNote command.

# SEE ALSO
Start-Process

