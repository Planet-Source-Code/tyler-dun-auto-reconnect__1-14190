<div align="center">

## DUN Auto Reconnect


</div>

### Description

Hits the Reconnect button when it finds the Reconnect Button. I have more to this code, but this is one part of it that is important.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tyler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tyler.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tyler-dun-auto-reconnect__1-14190/archive/master.zip)

### API Declarations

```
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOW = 5
```


### Source Code

```
'' Better Off Putting it in a Timer.
'' Set the Interval to 3000.
'' Private Sub Timer1_Timer()
dim findwin as Long
findwin = FindWindow("#32770", "Reestablish Connection")
If findwin <> 0 Then
Call ShowWindow(findwin, SW_SHOW)
SendKeys "{enter}", True
End If
'' End Sub
```

