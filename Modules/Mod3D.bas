Attribute VB_Name = "Mod3D"
'---------------------------------------------------------------------------------------
' Module    : Mod3D
' Author    : YPN
' Date      : 2018-03-24 00:17
' Purpose   : 3D控件用
'---------------------------------------------------------------------------------------

Option Explicit

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'  RaiseEvent MouseMove(Button, Shift, x, y)
'    If (x < 0) Or (y < 0) Or (x > .Width) Or (y > .Height) Then
'      ReleaseCapture
'      RaiseEvent MouseExit  ' 鼠标离开的代码
'    Else
'      SetCapture hWnd
'      RaiseEvent MouseIn   ' 鼠标进入的代码
'    End If
'End Sub


'***************************************************************************************************
'  '这个是调用复制 父级/屏幕背景的方法，在要透明的地方调用本函数
'***************************************************************************************************
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_ERASEBKGND = &H14
Private Const WM_PAINT = &HF
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
'***************************************************************************************************
'  '这个是调用复制 父级/屏幕背景的方法，在要透明的地方调用本函数
'***************************************************************************************************




' *************************************
' *            CONSTANTS              *
' *************************************
Private Const API_DIB_RGB_COLORS As Long = 0



' *************************************
' *        TYPES                      *
' *************************************
Public Type tpAPI_RECT                  ' NEVER ever use 'Left' or 'Right' as names in a udt!
    lLeft       As Long                 ' You run into trouble with the VB build-in functions for
    lTop        As Long                 ' string/variant handling (Left() and Right(). And this
    lRight      As Long                 ' strange effects and error messages are really hard to debug ... ;(
    lBottom     As Long
End Type

Private Type tpBITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type


' *************************************
' *        API DECLARES               *
' *************************************
Private Declare Function API_StretchDIBits Lib "gdi32" Alias "StretchDIBits" _
        (ByVal hDC As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal dx As Long, _
         ByVal dy As Long, _
         ByVal SrcX As Long, _
         ByVal SrcY As Long, _
         ByVal wSrcWidth As Long, _
         ByVal wSrcHeight As Long, _
         lpBits As Any, _
         lpBitsInfo As tpBITMAPINFOHEADER, _
         ByVal wUsage As Long, _
         ByVal dwRop As Long) As Long
'
'
'

'复制控件在父窗口的背景
Public Sub CopyBackGround(ByVal phWnd As Long, ByVal chWnd As Long, ByVal hDestDC As Long)
    Dim lpRect As RECT, lpPoint As POINTAPI, nWidth As Long, nHeight As Long
    Dim BitMap As Long, oldBitMap As Long, hDC As Long, memDC As Long
    
    Call GetWindowRect(phWnd, lpRect) 'Call GetClientRect(phWnd, lpRect)
    nWidth = lpRect.Right - lpRect.Left '获取控件的宽度
    nHeight = lpRect.Bottom - lpRect.Top '获取控件的高度
    
    hDC = GetDC(0)
    BitMap = CreateCompatibleBitmap(hDC, nWidth, nHeight)
    Call ReleaseDC(0, hDC)
    memDC = CreateCompatibleDC(0)
    oldBitMap = SelectObject(memDC, BitMap)
    Call SendMessage(phWnd, WM_ERASEBKGND, memDC, 0)
    Call SendMessage(phWnd, WM_PAINT, memDC, 0)
    '至此memDC上已经保存了父窗口的背景内容
    '用户可以调用BitBlt(...)等函数拷贝memDC的内容到子窗口的某个区域，
    '这样就达到了透明效果;
    Call GetWindowRect(chWnd, lpRect)
    lpPoint.X = lpRect.Left
    lpPoint.Y = lpRect.Top
    Call ScreenToClient(phWnd, lpPoint) '获取控件在父窗口的左上角位置
    Call BitBlt(hDestDC, 0, 0, nWidth, nHeight, memDC, lpPoint.X, lpPoint.Y, SRCCOPY)
    Call SelectObject(memDC, oldBitMap)
    Call DeleteDC(memDC)
    Call DeleteObject(BitMap)
    'UserControl.Refresh  '调用本过程后必须刷新★★★
End Sub
'复制控件在屏幕的背景
Public Sub CopyScreenBackground(ByVal phWnd As Long, ByVal chWnd As Long, ByVal hDestDC As Long)
    Dim lpRect As RECT, nWidth As Long, nHeight As Long, hDC As Long
    
    Call GetWindowRect(chWnd, lpRect)
    nWidth = lpRect.Right - lpRect.Left '获取控件的宽度
    nHeight = lpRect.Bottom - lpRect.Top '获取控件的高度
    
    ShowWindow chWnd, 0 '隐藏
    DoEvents
    hDC = GetDC(0)
    Call BitBlt(hDestDC, 0, 0, nWidth, nHeight, hDC, lpRect.Left, lpRect.Top, SRCCOPY)
    Call ReleaseDC(0, hDC)
    ShowWindow chWnd, 5 '显示
End Sub
'***************************************************************************************************
'  '这个是调用复制 父级/屏幕背景的方法，在要透明的地方调用本函数
'    If copyscreen = False Then
'        CopyScreenBackground UserControl.Parent.hwnd, UserControl.hwnd, UserControl.hDC
'    Else
'        CopyBackGround UserControl.Parent.hwnd, UserControl.hwnd, UserControl.hDC
'    End If
'***************************************************************************************************


Public Sub DrawTopDownGradient(hDC As Long, rc As tpAPI_RECT, ByVal lRGBColorFrom As Long, ByVal lRGBColorTo As Long)
    
    Dim uBIH            As tpBITMAPINFOHEADER
    Dim lBits()         As Long
    Dim lColor          As Long
    
    Dim X               As Long
    Dim Y               As Long
    Dim xEnd            As Long
    Dim yEnd            As Long
    Dim ScanlineWidth   As Long
    Dim yOffset         As Long
    
    Dim R               As Long
    Dim G               As Long
    Dim B               As Long
    Dim end_R           As Long
    Dim end_G           As Long
    Dim end_B           As Long
    Dim dR              As Long
    Dim dG              As Long
    Dim dB              As Long
    
    
    ' Split a RGB long value into components - FROM gradient color
    lRGBColorFrom = lRGBColorFrom And &HFFFFFF                      ' "SplitRGB"  by www.Abstractvb.com
    R = lRGBColorFrom Mod &H100&                                    ' Should be the fastest way in pur VB
    lRGBColorFrom = lRGBColorFrom \ &H100&                          ' See test on VBSpeed (http://www.xbeat.net/vbspeed/)
    G = lRGBColorFrom Mod &H100&                                    ' Btw: API solution with RTLMoveMem is slower ... ;)
    lRGBColorFrom = lRGBColorFrom \ &H100&
    B = lRGBColorFrom Mod &H100&
    
    ' Split a RGB long value into components - TO gradient color
    lRGBColorTo = lRGBColorTo And &HFFFFFF
    end_R = lRGBColorTo Mod &H100&
    lRGBColorTo = lRGBColorTo \ &H100&
    end_G = lRGBColorTo Mod &H100&
    lRGBColorTo = lRGBColorTo \ &H100&
    end_B = lRGBColorTo Mod &H100&
    
    
    '-- Loops bounds
    xEnd = rc.lRight - rc.lLeft
    yEnd = rc.lBottom - rc.lTop
    
    ' Check:  Top lower than Bottom ?
    If yEnd < 1 Then
    
        Exit Sub
    End If
    
    '-- Scanline width
    ScanlineWidth = xEnd + 1
    yOffset = -ScanlineWidth
    
    '-- Initialize array size
    ReDim lBits((xEnd + 1) * (yEnd + 1) - 1) As Long
       
    '-- Get color distances
    dR = end_R - R
    dG = end_G - G
    dB = end_B - B
       
    '-- Gradient loop over rectangle
    For Y = 0 To yEnd
        
        '-- Calculate color and *y* offset
        lColor = B + (dB * Y) \ yEnd + 256 * (G + (dG * Y) \ yEnd) + 65536 * (R + (dR * Y) \ yEnd)
        
        yOffset = yOffset + ScanlineWidth
        
        '-- *Fill* line
        For X = yOffset To xEnd + yOffset
            lBits(X) = lColor
        Next X
        
    Next Y
    
    '-- Prepare bitmap info structure
    With uBIH
        .biSize = Len(uBIH)
        .biBitCount = 32
        .biPlanes = 1
        .biWidth = xEnd + 1
        .biHeight = -yEnd + 1
    End With
    
    '-- Finaly, paint *bits* onto given DC
    API_StretchDIBits hDC, _
            rc.lLeft, rc.lTop, _
            xEnd, yEnd, _
            0, 0, _
            xEnd, yEnd, _
            lBits(0), _
            uBIH, _
            API_DIB_RGB_COLORS, _
            vbSrcCopy
End Sub
' #*#
