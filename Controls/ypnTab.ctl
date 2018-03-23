VERSION 5.00
Begin VB.UserControl ypnTab 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF00DA&
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   ControlContainer=   -1  'True
   MaskColor       =   &H00FF00DA&
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ToolboxBitmap   =   "ypnTab.ctx":0000
   Begin VB.PictureBox skin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   4320
      Picture         =   "ypnTab.ctx":0312
      ScaleHeight     =   2400
      ScaleWidth      =   1800
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox skinPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00DA&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   720
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   2685
   End
End
Attribute VB_Name = "ypnTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4

Enum skinCount
    skin0 = 0
    skin1 = 1
    skin2 = 2
    skin3 = 3
    skin4 = 4
    skin5 = 5
    skin6 = 6
    skin7 = 7
End Enum
Enum activeTab
    tab01 = 0
    tab02 = 1
    tab03 = 2
    tab04 = 3
    tab05 = 4
    tab06 = 5
    tab07 = 6
    tab08 = 7
    tab09 = 8
    tab10 = 9
    tab11 = 10
    tab12 = 11
    tab13 = 12
    tab14 = 13
    tab15 = 14
    tab16 = 15
    tab17 = 16
    tab18 = 17
    tab19 = 18
    tab20 = 19
    tab21 = 20
    tab22 = 21
    tab23 = 22
    tab24 = 23
    tab25 = 24
    tab26 = 25
    tab27 = 26
    tab28 = 27
    tab29 = 28
    tab30 = 29
End Enum

Dim Width As Single, Height As Single
Const BlankSpace = 5  '名字距标签边界距离
Const ControlSpace = 63000

Dim ClabelActHeight As Integer   '激活标签高度
Dim ClabelHeight As Integer      '未激活标签高度
Dim maxHeight As Integer
Dim ClabelWidth As Integer
Dim ClabelBlank As Integer       '标签间隔
Dim labelWholeWidth As Integer

Dim myLabel() As String
Dim labelCount As Integer

Dim labelActive As activeTab
Dim labelAim As Integer


Dim ControlSkinIndex As skinCount
Dim TabBackColor As Long
Dim TabEnabled As Boolean
Dim lineColor(7) As Long
Dim TabColor(7) As Long
Dim captionRect As RECT
Dim setCaptureAPI As Boolean
Public Event TabSwitch(ByVal LastActiveTab As Integer)

Public Sub Refresh()
    Debug.Print "bg refresh" & Timer
    UserControl.BackStyle = 1
    Dim tempWidth As Single
    UserControl.Cls
    'UserControl.BackColor = TabBackColor
    UserControl.Line (0, maxHeight - 1)-(Width - 1, Height - 1), TabBackColor, BF
    UserControl.Line (0, maxHeight - 1)-(Width - 1, Height - 1), lineColor(ControlSkinIndex), B
    For i = 0 To labelCount - 1
        tempWidth = i * (ClabelWidth + ClabelBlank) + ClabelBlank
        captionRect.Left = tempWidth + BlankSpace
        captionRect.Right = tempWidth + ClabelWidth - BlankSpace
        captionRect.Bottom = maxHeight - BlankSpace
        If labelActive = i Then
            captionRect.Top = maxHeight - ClabelActHeight + BlankSpace
            UserControl.PaintPicture skinPicture.Image, tempWidth, 0, ClabelWidth, maxHeight, 0, 0, ClabelWidth, maxHeight, vbSrcCopy
        ElseIf labelAim = i Then
            captionRect.Top = maxHeight - ClabelHeight + BlankSpace
            UserControl.PaintPicture skinPicture.Image, tempWidth, 0, ClabelWidth, maxHeight, ClabelWidth, 0, ClabelWidth, maxHeight, vbSrcCopy
        Else
            captionRect.Top = maxHeight - ClabelHeight + BlankSpace
            UserControl.PaintPicture skinPicture.Image, tempWidth, 0, ClabelWidth, maxHeight, ClabelWidth * 2, 0, ClabelWidth, maxHeight, vbSrcCopy
        End If
        UserControl.ForeColor = IIf(TabEnabled, &H0, &H99A8AC)
        DrawText UserControl.hDC, myLabel(i), -1, captionRect, DT_CENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_VCENTER
    Next
    Set UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0
    Debug.Print "ed refresh" & Timer
End Sub
Public Sub refreshSkin()
    Debug.Print "bg refreshskin" & Timer
    'skinPicture.BackColor = TabBackColor
    skinPicture.Cls
    skinPicture.Width = ClabelWidth * 3
    maxHeight = IIf(ClabelActHeight > ClabelHeight, ClabelActHeight, ClabelHeight)
    skinPicture.Height = maxHeight
    Call drawSkin(0, ClabelActHeight, 0) '激活
    Call drawSkin(ClabelWidth, ClabelHeight, 1)
    Call drawSkin(ClabelWidth + ClabelWidth, ClabelHeight, 2)
    Call Refresh
    Debug.Print "ed refreshskin" & Timer
End Sub
Private Sub drawSkin(Left As Integer, TempHeight As Integer, Condition As Integer) 'PaintPicture Pic,destX,destY,destWidth,destHeight,scrX,scrY,scrWidth,scrHeight
    Debug.Print "bg drawSkin" & Timer
    Dim srcLeft As Single, Top As Single
    srcLeft = Condition * 40
    Top = maxHeight - TempHeight
    
    skinPicture.PaintPicture skin.Image, Left, Top, 5, 5, srcLeft, ControlSkinIndex * 20 + 0, 5, 5, vbSrcCopy
    skinPicture.PaintPicture skin.Image, Left, Top + 5, 5, TempHeight - 10, srcLeft, ControlSkinIndex * 20 + 5, 5, 10, vbSrcCopy
    skinPicture.PaintPicture skin.Image, Left, Top + TempHeight - 5, 5, 5, srcLeft, ControlSkinIndex * 20 + 15, 5, 5, vbSrcCopy
    
    skinPicture.PaintPicture skin.Image, Left + 5, Top, ClabelWidth - 10, 5, srcLeft + 5, ControlSkinIndex * 20 + 0, 30, 5, vbSrcCopy
    skinPicture.PaintPicture skin.Image, Left + 5, Top + 5, ClabelWidth - 10, TempHeight - 10, srcLeft + 5, ControlSkinIndex * 20 + 5, 30, 10, vbSrcCopy
    skinPicture.PaintPicture skin.Image, Left + 5, Top + TempHeight - 5, ClabelWidth - 10, 5, srcLeft + 5, ControlSkinIndex * 20 + 15, 30, 5, vbSrcCopy
    
    skinPicture.PaintPicture skin.Image, Left + ClabelWidth - 5, Top, 5, 5, srcLeft + 35, ControlSkinIndex * 20 + 0, 5, 5, vbSrcCopy
    skinPicture.PaintPicture skin.Image, Left + ClabelWidth - 5, Top + 5, 5, TempHeight - 10, srcLeft + 35, ControlSkinIndex * 20 + 5, 5, 10, vbSrcCopy
    skinPicture.PaintPicture skin.Image, Left + ClabelWidth - 5, Top + TempHeight - 5, 5, 5, srcLeft + 35, ControlSkinIndex * 20 + 15, 5, 5, vbSrcCopy
    Call drawBackColor(skinPicture.hDC, CLng(Left), CLng(Top))
    Call drawBackColor(skinPicture.hDC, CLng(Left + ClabelWidth - 5), CLng(Top))
    Debug.Print "ed drawSkin" & Timer
End Sub
Private Sub drawBackColor(hDC As Long, Left As Long, Top As Long)
    Debug.Print "bg drawBackColor" & Timer
    Dim i As Long, j As Long
    For i = 0 To 5
        For j = 0 To 5
            If GetPixel(hDC, Left + i, Top + j) = &HFF00FF Then SetPixel hDC, Left + i, Top + j, &HFF00DA
        Next
    Next
    Debug.Print "ed drawBackColor" & Timer
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempLabel As Integer
    If Not TabEnabled Then Exit Sub
    labelAim = -1
    If Y > 0 And Y < ClabelHeight And X > 0 And X < labelWholeWidth Then
        tempLabel = X \ (ClabelWidth + ClabelBlank)
        If (X Mod (ClabelWidth + ClabelBlank)) > ClabelBlank Then
            labelAim = tempLabel
            If Not setCaptureAPI Then
                Call SetCapture(UserControl.hwnd)
                setCaptureAPI = True
            End If
        End If
    Else
        If setCaptureAPI Then
            Call ReleaseCapture
            setCaptureAPI = False
        End If
    End If
    'Call Refresh
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, Left As Integer, lastActive As Integer
    If Not TabEnabled Then Exit Sub
    If labelAim = -1 Or labelAim = labelActive Then Exit Sub
    lastActive = labelActive
    labelActive = labelAim
    Call ControlRefresh(lastActive)
    Call Refresh
    RaiseEvent TabSwitch(lastActive)
End Sub
Private Sub ControlRefresh(lastActive As Integer)
    On Error Resume Next
    Dim Ctl As Control, MoveVal&, oldtop&, oldTabstop&
    If lastActive = labelActive Then Exit Sub
    MoveVal = (lastActive - labelActive)
    For Each Ctl In UserControl.ContainedControls
        err.Number = 0
        If Ctl.Top >= ControlSpace Then
            oldtop = Ctl.Top - ControlSpace
        Else
            oldtop = Ctl.Top
        End If
        Ctl.WhatsThisHelpID = Val(Ctl.WhatsThisHelpID) + MoveVal
        If Ctl.WhatsThisHelpID = 1000 Or Ctl.WhatsThisHelpID = 100000 Then
            If err.Number = 0 Then Ctl.Top = oldtop
        Else
            If err.Number = 0 Then Ctl.Top = ControlSpace + oldtop
            If Ctl.TabStop = True Then
                If Ctl.WhatsThisHelpID > 99000 Then
                    Ctl.WhatsThisHelpID = Ctl.WhatsThisHelpID + 1000 - 100000
                Else
                    Ctl.WhatsThisHelpID = Ctl.WhatsThisHelpID + 1000
                End If
            ElseIf ctltabstop = False And Ctl.WhatsThisHelpID < 500 Then
                Ctl.WhatsThisHelpID = Ctl.WhatsThisHelpID + 100000
            End If
            Ctl.TabStop = False
            
        End If
        
        If Ctl.WhatsThisHelpID = 1000 Then
            Ctl.TabStop = True
            Ctl.WhatsThisHelpID = 0
        End If
    Next Ctl
End Sub

Public Property Get Caption() As String
    Caption = myLabel(labelActive)
End Property
Public Property Let Caption(ByVal vNewValue As String)
    myLabel(labelActive) = vNewValue
    Call Refresh
End Property
Public Property Get TabCount() As Integer
    TabCount = labelCount
End Property
Public Property Let TabCount(ByVal vNewValue As Integer)
    Dim i As Integer, lastActive As Integer
    If labelCount = vNewValue Then Exit Property
    labelCount = vNewValue
    ReDim Preserve myLabel(labelCount - 1)
    If labelActive > labelCount - 1 Then
        labelActive = labelCount - 1
        RaiseEvent TabSwitch(lastActive)
    End If
    Call Refresh
    labelWholeWidth = (ClabelWidth + ClabelBlank) * labelCount
End Property

Public Property Get activeTab() As activeTab
    activeTab = labelActive
End Property

Public Property Let activeTab(ByVal vNewValue As activeTab)
    Dim lastActive As Integer
    If vNewValue < 0 Or vNewValue >= labelCount Then vNewValue = 0
    
    If labelActive <> vNewValue Then
        lastActive = labelActive
        labelActive = vNewValue
        Call ControlRefresh(lastActive)
        Call Refresh
        RaiseEvent TabSwitch(lastActive)
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = TabBackColor
End Property
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    TabBackColor = vNewValue
    UserControl.BackColor = vNewValue
    Call refreshSkin
End Property
Public Property Get skinIndex() As skinCount
    skinIndex = ControlSkinIndex
End Property
Public Property Let skinIndex(ByVal vNewValue As skinCount)
    ControlSkinIndex = vNewValue
    Call refreshSkin
End Property
Public Property Get LabelActHeight() As Integer
    LabelActHeight = ClabelActHeight
End Property
Public Property Let LabelActHeight(ByVal vNewValue As Integer)
    If vNewValue < 15 Then vNewValue = 15
    ClabelActHeight = vNewValue
    Call refreshSkin
End Property
Public Property Get LabelHeight() As Integer
    LabelHeight = ClabelHeight
End Property
Public Property Let LabelHeight(ByVal vNewValue As Integer)
    If vNewValue < 15 Then vNewValue = 15
    ClabelHeight = vNewValue
    Call refreshSkin
End Property
Public Property Get labelWidth() As Integer
    labelWidth = ClabelWidth
End Property
Public Property Let labelWidth(ByVal vNewValue As Integer)
    If vNewValue < 15 Then vNewValue = 15
    ClabelWidth = vNewValue
    labelWholeWidth = (ClabelWidth + ClabelBlank) * labelCount
    Call refreshSkin
End Property
Public Property Get LabelBlank() As Integer
    LabelBlank = ClabelBlank
End Property
Public Property Let LabelBlank(ByVal vNewValue As Integer)
    If vNewValue < 0 Then Exit Property
    ClabelBlank = vNewValue
    labelWholeWidth = (ClabelWidth + ClabelBlank) * labelCount
    Call Refresh
End Property
Public Property Get Enabled() As Boolean
    Enabled = TabEnabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
    If TabEnabled <> vNewValue Then
        TabEnabled = vNewValue
        Call Refresh
    End If
End Property
Private Sub UserControl_Initialize()
    Dim i As Integer
    For i = 0 To 7
        lineColor(i) = GetPixel(skin.hDC, 119, i * 20 + 19)
        TabColor(i) = GetPixel(skin.hDC, 20, i * 20 + 19)
    Next
    TabBackColor = GetSysColor(15) ' &HE0DFE3
    ClabelActHeight = 29
    ClabelHeight = 24
    ClabelWidth = 80
    ClabelBlank = 2
    TabEnabled = True
    ControlSkinIndex = skin0
    
    labelCount = 3
    ReDim myLabel(labelCount - 1)
    myLabel(0) = "新标签1"
    myLabel(1) = "新标签2"
    myLabel(2) = "新标签3"
    labelActive = 0
    labelAim = -1
End Sub
Private Sub UserControl_Resize()
    If UserControl.ScaleWidth < 80 Then UserControl.Width = 80 * 15
    If UserControl.ScaleHeight < 30 Then UserControl.Height = 30 * 15
    Width = UserControl.ScaleWidth
    Height = UserControl.ScaleHeight
    Call refreshSkin
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim i As Integer
    TabCount = PropBag.ReadProperty("TabCount", 3)
    activeTab = PropBag.ReadProperty("ActiveTab", 0)
    TabEnabled = PropBag.ReadProperty("enabled", True)
    TabBackColor = PropBag.ReadProperty("BackColor", GetSysColor(15))
    skinIndex = PropBag.ReadProperty("skinIndex", 0)
    LabelActHeight = PropBag.ReadProperty("LabelActHeight", 29)
    LabelHeight = PropBag.ReadProperty("LabelHeight", 24)
    labelWidth = PropBag.ReadProperty("LabelWidth", 80)
    LabelBlank = PropBag.ReadProperty("LabelBlank", 2)
    ReDim myLabel(TabCount - 1)
    For i = 0 To TabCount - 1
        myLabel(i) = PropBag.ReadProperty("Caption" + CStr(i), "新标签")
    Next
    Call ControlRefresh(0)
    Call refreshSkin
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim i As Integer
    Call PropBag.WriteProperty("TabCount", TabCount, 3)
    Call PropBag.WriteProperty("ActiveTab", activeTab, 0)
    Call PropBag.WriteProperty("enabled", TabEnabled, True)
    Call PropBag.WriteProperty("BackColor", TabBackColor, 0)
    Call PropBag.WriteProperty("skinIndex", skinIndex, 0)
    Call PropBag.WriteProperty("LabelActHeight", LabelActHeight, 29)
    Call PropBag.WriteProperty("LabelHeight", LabelHeight, 24)
    Call PropBag.WriteProperty("LabelWidth", labelWidth, 80)
    Call PropBag.WriteProperty("LabelBlank", LabelBlank, 2)
    For i = 0 To TabCount - 1
        Call PropBag.WriteProperty("Caption" + CStr(i), myLabel(i), "新标签")
    Next
End Sub
