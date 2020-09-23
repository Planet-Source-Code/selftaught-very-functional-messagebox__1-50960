Attribute VB_Name = "mLavolpe"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal Hwnd As Long, ByVal lpString As String) As Long

Public Function lv_TimerCallBack(ByVal Hwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tgtButton As lvButton
    CopyMemory tgtButton, GetProp(Hwnd, "lv_ClassID"), &H4
    Call tgtButton.TimerUpdate(GetProp(Hwnd, "lv_TimerID"))  ' fire the button's event
    CopyMemory tgtButton, 0&, &H4                                    ' erase this instance
End Function
