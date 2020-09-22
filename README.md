<div align="center">

## Always On Top


</div>

### Description

Keeps Your Form On Top

This Is Really Kewl Because You Can Just Use A false attribuite to set it as not on top instead of using 2 functions
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Jacob Muller](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-jacob-muller.md)
**Level**          |Unknown
**User Rating**    |4.3 (51 globes from 12 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-jacob-muller-always-on-top__1-2651/archive/master.zip)

### API Declarations

```
Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal X As Long, _
  ByVal Y As Long, _
  ByVal cx As Long, _
  ByVal cy As Long, _
  ByVal wFlags As Long) As Long
```


### Source Code

```
Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
  If SetOnTop Then
   lFlag = HWND_TOPMOST
  Else
   lFlag = HWND_NOTOPMOST
  End If
  SetWindowPos myfrm.hwnd, lFlag, _
  myfrm.Left / Screen.TwipsPerPixelX, _
  myfrm.Top / Screen.TwipsPerPixelY, _
  myfrm.Width / Screen.TwipsPerPixelX, _
  myfrm.Height / Screen.TwipsPerPixelY, _
  SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
'Well, if your for example in a form called 'Form1' then you'd simply type:
AlwaysOnTop Form1, True
```

