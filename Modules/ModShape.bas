Attribute VB_Name = "ModShape"
'---------------------------------------------------------------------------------------
' Module    : ModShape
' Author    : YPN
' Date      : 2018-03-24 00:17
' Purpose   : 变形控件用
'---------------------------------------------------------------------------------------

Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Function lv_TimerCallBack(ByVal hwnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim tgtButton As ypnButton_Shape
    ' when timer was intialized, the button control's hWnd
    ' had property SET to the handle of the control itself
    ' AND the timer ID was also SET as a window property
    CopyMemory tgtButton, GetProp(hwnd, "lv_ClassID"), &H4
    Call tgtButton.TimerUpdate(GetProp(hwnd, "lv_TimerID"))  ' fire the button's event
    CopyMemory tgtButton, 0&, &H4
    ' erase this instance
End Function

