---
external help file: StickyNotes-help.xml
schema: 2.0.0
online version: 
---

# Remove-StickyNote
## SYNOPSIS
Remove a new sticky note.
## SYNTAX

```
Remove-StickyNote [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This command will remove the sticky note that has focus.
If you have more than one note, the last one created has focus, unless you manually select a different one.

If you kill the 'StikyNot.exe' process, the next time you create a note, any previously created notes will return.
## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> remove-stickynote
```

## PARAMETERS

### -WhatIf
```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -Confirm
```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).
## INPUTS

### None

## OUTPUTS

### None

## NOTES
Version     : 1.2

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/
## RELATED LINKS

[New-StickyNote]()
