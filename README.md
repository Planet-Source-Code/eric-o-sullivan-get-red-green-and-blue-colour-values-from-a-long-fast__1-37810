<div align="center">

## Get Red, Green and Blue colour values from a Long \- FAST


</div>

### Description

This procedure takes a Long colour value and splits it into it's component Red, Green and Blue values.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eric O'Sullivan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eric-o-sullivan.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eric-o-sullivan-get-red-green-and-blue-colour-values-from-a-long-fast__1-37810/archive/master.zip)





### Source Code

```
Public Sub GetRGB(ByVal lngColour As Long, _
     Optional ByRef intRed As Integer, _
     Optional ByRef intGreen As Integer, _
     Optional ByRef intBlue As Integer)
 'Convert Long to RGB:
          '  Sys Cols, Blue , Green , Red
 Const RED_MASK As Long = 255  'Binary: 00000000,00000000,00000000,11111111
 Const GREEN_MASK As Long = 65280 'Binary: 00000000,00000000,11111111,00000000
 Const BLUE_MASK As Long = 16711680 'Binary: 00000000,11111111,00000000,00000000
 Const MAX_COLOUR As Long = 16777216 'Binary: 01111111,11111111,11111111,11111111
 'if the lngcolour value is greater than acceptable
 'then half the value
 If (lngColour > MAX_COLOUR) Or _
  (lngColour < -MAX_COLOUR) Then
  'system colour range
  Exit Sub
 End If
 'get the red, green and blue values
 intRed = (lngColour And RED_MASK)
 intGreen = (lngColour And GREEN_MASK) / 256 'shift 8 bits to the right
 intBlue = (lngColour And BLUE_MASK) / 65536 'shift 16 bits to the right
End Sub
```

