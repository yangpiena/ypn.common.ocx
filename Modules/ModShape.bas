Attribute VB_Name = "ModShape"
'***************************************************
'**系统名称：菲麦森装备制造业产前数据准备系统
'**模块描述：用户控件按钮用到的
'**模 块 名：ModLvTimer
'**创 建 人：杨东
'**日    期：2014-09-10 15:14:22
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V1.0.0
'***************************************************
' THIS MODULE WAS NOT WRITTEN BY DEAN CAMERA. I CANNOT OFFER ANY SUPPORT FOR THIS MODULE.


' REQUIRED: copy & paste these few lines in any module of your project
' This is used by every lvButtons control as a replacement of the Timer control
' By doing it this way, each button control does NOT need an individual timer control
' The timer function is primarily used to determine when the mouse enters/leaves
' the button's physical region on the screen.

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

