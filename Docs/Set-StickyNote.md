---
external help file: StickyNotes-help.xml
online version: 
schema: 2.0.0
---

# Set-StickyNote
## SYNOPSIS
Set text and style for a sticky note.

## SYNTAX

```
Set-StickyNote [[-Text] <String>] [-Bold] [-Italic] [-Underline] [-Alignment <String>] [-Append] [-FontSize <Int32>]
```

## DESCRIPTION
This command will set the text and style for a sticky note.
If there are multiple notes, this command will set the one that has focus, which you must do manually.
The default is the last note created.

The style parameters like -Bold behave more like toggles.
If text is already in bold than using -Bold will turn it off and vice versa.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
PS C:\> set-stickynote "pickup bread on the way home" -bold -underline -append
```

## PARAMETERS

### -Text
The text for the sticky note.

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 1
Default value: 
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Set the text to bold faced.

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
Set the text to italic.

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
Set the text to underline.

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
Set the text to the specified alignment. Valid values are Left, Center or Right.

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

### -Append
Append your text to the end of the note.
Otherwise existing text will be replaced

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

### -FontSize
Adjust the font size up or down. Use a positive number to increase
and a negative number to decrease. A value of 0 will not do 
anything.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: 

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

## INPUTS

## OUTPUTS

## NOTES
Version     : 1.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

## RELATED LINKS
[New-StickyNote]()
