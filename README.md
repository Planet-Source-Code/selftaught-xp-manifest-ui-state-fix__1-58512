<div align="center">

## XP Manifest UI State fix


</div>

### Description

Perhaps you've noticed that the VB Form engine does not support dynamic changing of the ui state. It gets the ui state when the form is first created and then never again. Obviously this defeats the whole purpose of it. For example, if the user opens your form using the mouse then the focus rectangles will not show on the command buttons or on other controls that conform to XP ui standards. This code a kludge that will show all ui states instead of only the state that was active when the form was created. Just call this sub in the initialize event of a form that is linked to CC 6 with a manifest.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[selftaught](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/selftaught.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/selftaught-xp-manifest-ui-state-fix__1-58512/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
```


### Source Code

```
Public Sub ForceWindowToShowAllUIStates(ByVal hwnd As Long)
  Const WM_CHANGEUISTATE As Long = &H127
  Const UIS_SET As Long = 1
  Const UIS_CLEAR As Long = 2
  Const UISF_HIDEACCEL As Long = &H2
  Const UISF_HIDEFOCUS As Long = &H1
  Const CLEAR_IT_ALL As Long = ((UISF_HIDEACCEL Or UISF_HIDEFOCUS) * &H10000) Or UIS_CLEAR
  SendMessage hwnd, WM_CHANGEUISTATE, CLEAR_IT_ALL, 0&
  SendMessage hwnd, WM_CHANGEUISTATE, UIS_SET, 0&
End Sub
```

