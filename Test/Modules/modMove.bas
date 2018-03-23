Attribute VB_Name = "modMOVE"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub DoDrag(TheForm As Form)
    If TheForm.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage TheForm.hwnd, &HA1, 2, 0&
    End If
End Sub
