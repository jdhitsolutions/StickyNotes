---
external help file: StickyNotes-help.xml
schema: 2.0.0
online version: 
---

# New-StickyNote
## SYNOPSIS
Create a new sticky note.
## SYNTAX

```
New-StickyNote [-Text] <String> [-Bold] [-Italic] [-Underline] [-Alignment <String>] [-FontFamily <String>]
 [<CommonParameters>]
```

## DESCRIPTION
This command will create a sticky note with the specified text
and formatting options.
## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> New-StickyNote "pickup milk on the way home" -bold -alignment center
```

Create a sticky note with bold text and centered
### -------------------------- EXAMPLE 2 --------------------------
```
PS C:\> (get-date).ToShortDateString() | new-stickynote
```

Pipe the current date as a string to a new sticky note.
## PARAMETERS

### -Text
Enter text for the sticky note.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: True
Position: 1
Default value: 
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Bold
Make text bold faced.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Make text italic.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Underline text.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Alignment
Set the text alignment. Possible values for Left, Center and Right.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: Left
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
Enter the name of a font family like Consolas, Verdana or Tahoma.
You can see a list of names with this command:

    [System.Drawing.Text.InstalledFontCollection]::new().Families.Name

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

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
Version     : 1.2.0.3

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/
## RELATED LINKS

[Set-StickyNote]()
[Remove-StickyNote]()
