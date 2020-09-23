<div align="center">

## Real Window Snap to


</div>

### Description

This code produces a snap to effect exactly like Winamp.

Uses POINTAPI type and GetCursorPos API.

It gets the current x and y does a few calculations and snaps-to the screen edge.

It does take into account for the taskbar but that may need some tweaking.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[DoWnLoHo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/downloho.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/downloho-real-window-snap-to__1-8604/archive/master.zip)





### Source Code

```
'Note** This meant to be saved as a form
'Copy below this line; paste into notepad; Save as frmSnapto.frm
VERSION 5.00
Begin VB.Form frmSnapTo
  BorderStyle   =  0 'None
  Caption     =  "Form1"
  ClientHeight  =  1335
  ClientLeft   =  0
  ClientTop    =  0
  ClientWidth   =  3660
  LinkTopic    =  "Form1"
  ScaleHeight   =  1335
  ScaleWidth   =  3660
  ShowInTaskbar  =  0  'False
  StartUpPosition =  3 'Windows Default
  Begin VB.Timer tmrPos
   Enabled     =  0  'False
   Interval    =  1
   Left      =  120
   Top       =  360
  End
  Begin VB.Label lblTop
   BackColor    =  &H000000FF&
   Caption     =  "Caption"
   Height     =  255
   Left      =  0
   TabIndex    =  0
   Top       =  0
   Width      =  3720
  End
End
Attribute VB_Name = "frmSnapTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Dim iX As Integer, iY As Integer
Private Sub lblTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
iX% = X: iY% = Y
tmrPos.Enabled = True
End Sub
Private Sub lblTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrPos.Enabled = False
End Sub
Private Sub tmrPos_Timer()
Dim ptPos As POINTAPI
 Call GetCursorPos(ptPos)
 lblTop.Caption = ptPos.X & " - " & ptPos.Y
If ptPos.Y - ((lblTop.Top + iY%) / Screen.TwipsPerPixelY) <= 20 Then ptPos.Y = 0 + ((lblTop.Top + iY%) / Screen.TwipsPerPixelY)
If ptPos.X - ((lblTop.Left + iX%) / Screen.TwipsPerPixelX) <= 20 Then ptPos.X = 0 + ((lblTop.Left + iX%) / Screen.TwipsPerPixelX)
If ptPos.Y - ((lblTop.Top + iY%) / Screen.TwipsPerPixelY) >= (Screen.Height - Me.Height - 400) / Screen.TwipsPerPixelY - 20 Then
  ptPos.Y = (Screen.Height - Me.Height + iY% - 400) / Screen.TwipsPerPixelY
End If
If ptPos.X - ((lblTop.Left + iX%) / Screen.TwipsPerPixelX) >= (Screen.Width - Me.Width) / Screen.TwipsPerPixelX - 20 Then
  ptPos.X = (Screen.Width - Me.Width + iX%) / Screen.TwipsPerPixelX
End If
Me.Top = (ptPos.Y * Screen.TwipsPerPixelY) - lblTop.Top - iY%
Me.Left = (ptPos.X * Screen.TwipsPerPixelX) - lblTop.Left - iX%
End Sub
```

