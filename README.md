<div align="center">

## Refresh Desktop


</div>

### Description

Refresh / Redraw / Repaint the users desktop. Used when you have changed something on the desktop and you want the user to see it :) - Comments and Rates will be appreciated.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[F2](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/f2.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/f2-refresh-desktop__1-34145/archive/master.zip)

### API Declarations

```
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" _
 Alias "FindWindowExA" (ByVal hWnd1 As Long, _
 ByVal hWnd2 As Long, ByVal lpsz1 As String, _
 ByVal lpsz2 As String) As Long
Const WM_KEYDOWN = &H100
```


### Source Code

```
Public Sub RefreshDesktop()
 Dim DeskWin As Long
 DeskWin = FindWindowEx(0&, 0&, "Progman", vbNullString)
 DeskWin = FindWindowEx(DeskWin, 0&, "SHELLDLL_DefView", vbNullString)
 DeskWin = FindWindowEx(DeskWin, 0&, "SysListView32", vbNullString)
 PostMessage DeskWin, WM_KEYDOWN, &H74, &H3F0001
End Sub
```

