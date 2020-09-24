<div align="center">

## Check if a file exist\!


</div>

### Description

This will check if a dam*n file exists the easy and "right" way! Plz vote for this code if you like it!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VisualBlind](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/visualblind.md)
**Level**          |Intermediate
**User Rating**    |3.1 (34 globes from 11 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/visualblind-check-if-a-file-exist__1-11919/archive/master.zip)





### Source Code

```
Dim fso
Dim strFile As String
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists("c:\windows\Calc.exe") Then
  MsgBox "It does exist", vbInformation, "Does Exist"
  Else
  MsgBox "It does not exist!", vbExclamation, "Doesn't Exist"
  End If
```

