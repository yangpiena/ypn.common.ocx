VERSION 5.00
Begin VB.UserControl ypnSwitch 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  '透明
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H000000FF&
   MaskPicture     =   "ypnSwitch.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ypnSwitch.ctx":A9E2
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Index           =   2
      Left            =   4200
      Picture         =   "ypnSwitch.ctx":ACF4
      ScaleHeight     =   1920
      ScaleWidth      =   540
      TabIndex        =   2
      Top             =   0
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2880
      Index           =   1
      Left            =   3120
      Picture         =   "ypnSwitch.ctx":E336
      ScaleHeight     =   2880
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   0
      Width           =   1110
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1800
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2760
      Index           =   0
      Left            =   1920
      Picture         =   "ypnSwitch.ctx":18B7A
      ScaleHeight     =   2760
      ScaleWidth      =   1170
      TabIndex        =   0
      Top             =   0
      Width           =   1170
   End
   Begin VB.Image MaskPic 
      Height          =   240
      Index           =   5
      Left            =   120
      Picture         =   "ypnSwitch.ctx":2355C
      Top             =   2760
      Width           =   540
   End
   Begin VB.Image MaskPic 
      Height          =   240
      Index           =   4
      Left            =   120
      Picture         =   "ypnSwitch.ctx":237F0
      Top             =   2640
      Width           =   540
   End
   Begin VB.Image MaskPic 
      Height          =   360
      Index           =   3
      Left            =   120
      Picture         =   "ypnSwitch.ctx":23A84
      Top             =   2400
      Width           =   1110
   End
   Begin VB.Image MaskPic 
      Height          =   360
      Index           =   2
      Left            =   120
      Picture         =   "ypnSwitch.ctx":241F8
      Top             =   2160
      Width           =   1110
   End
   Begin VB.Image MaskPic 
      Height          =   345
      Index           =   1
      Left            =   120
      Picture         =   "ypnSwitch.ctx":2496C
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Image MaskPic 
      Height          =   345
      Index           =   0
      Left            =   120
      Picture         =   "ypnSwitch.ctx":250F0
      Top             =   1920
      Width           =   1170
   End
End
Attribute VB_Name = "ypnSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ypnSwitch
' Author    : YPN
' Date      : 2018-03-24 00:15
' Purpose   : 扁平化的开关
'---------------------------------------------------------------------------------------

Dim SwitchEnable As Boolean '有效无效
Dim SwitchCondition As Integer '0普通1指向2按下
Dim SwitchValue As Boolean '启用关闭
Dim newValue As Boolean

Dim SwitchAim As Boolean '鼠标在控件上
Enum mySkin
    Safe = 0
    Kingsoft = 1
    KuGou = 2
End Enum
Dim SwitchSkin As mySkin

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          'Aki
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long
Private Type POINT_API
    X As Long
    Y As Long
End Type

Public Event Click(Value As Boolean)
Public Event MouseOut()
Public Event MouseIn()

Private Sub UserControl_Initialize()
    Call Refresh
End Sub
Private Sub UserControl_InitProperties()
    SwitchCondition = 0
    SwitchEnable = True
    SwitchValue = True
    SwitchAim = False
    SwitchSkin = Safe
End Sub
Private Sub Refresh() '0普通1指向2按下
    If SwitchSkin = Safe Then
        UserControl.PaintPicture Picture1(0).Picture, 0, 0, 78, 23, 0, SwitchCondition * 23, 78, 23
        Set UserControl.MaskPicture = MaskPic(IIf(SwitchValue, 0, 1)).Picture
    ElseIf SwitchSkin = Kingsoft Then
        UserControl.PaintPicture Picture1(1).Picture, 0, 0, 74, 24, 0, SwitchCondition * 24, 74, 24
        Set UserControl.MaskPicture = MaskPic(IIf(SwitchValue, 2, 3)).Picture
    Else
        UserControl.PaintPicture Picture1(2).Picture, 0, 0, 36, 16, 0, SwitchCondition * 16, 36, 16
        Set UserControl.MaskPicture = MaskPic(IIf(SwitchValue, 4, 5)).Picture
    End If
    UserControl.Refresh
End Sub
Private Sub ChangeSkin()
    If SwitchSkin = Safe Then
        UserControl.Width = 78 * 15
        UserControl.Height = 23 * 15
    ElseIf SwitchSkin = Kingsoft Then
        UserControl.Width = 74 * 15
        UserControl.Height = 24 * 15
    Else
        UserControl.Width = 36 * 15
        UserControl.Height = 16 * 15
    End If
    Call Refresh
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And SwitchEnable Then
        newValue = Not SwitchValue
        SwitchCondition = IIf(SwitchValue, 2, 6)
        Call Refresh
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And SwitchEnable And SwitchAim Then
        SwitchValue = newValue
        SwitchCondition = IIf(SwitchValue, 1, 5)
        Call Refresh
        RaiseEvent Click(SwitchValue)
    End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Timer1.Enabled Then
    End If
End Sub
Private Sub Timer1_Timer()
    If Not SwitchEnable Then Exit Sub
    Dim dot As POINT_API
    Call GetCursorPos(dot)
    ScreenToClient UserControl.hwnd, dot
    If dot.X < UserControl.ScaleLeft Or dot.Y < UserControl.ScaleTop Or _
        dot.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or dot.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        If SwitchAim Then '移出
            SwitchAim = False
            SwitchCondition = IIf(SwitchValue, 0, 4)
            Call Refresh
            RaiseEvent MouseOut
        End If
    Else
        If Not SwitchAim Then
            SwitchAim = True
            SwitchCondition = IIf(SwitchValue, 1, 5)
            Call Refresh
            RaiseEvent MouseIn
        End If
    End If
End Sub
Public Property Get Enable() As Boolean
    Enable = SwitchEnable
End Property
Public Property Let Enable(ByVal vNewValue As Boolean)
    SwitchEnable = vNewValue
    
    If SwitchEnable Then
        If SwitchAim Then
            SwitchCondition = IIf(SwitchValue, 1, 5)
        Else
            SwitchCondition = IIf(SwitchValue, 0, 4)
        End If
    Else
        SwitchCondition = IIf(SwitchValue, 3, 7)
    End If
    Call Refresh
End Property
Public Property Get Value() As Boolean
    Value = SwitchValue
End Property
Public Property Let Value(ByVal vNewValue As Boolean)
    SwitchValue = vNewValue
    
    If SwitchEnable Then
        If SwitchAim Then
            SwitchCondition = IIf(SwitchValue, 1, 5)
        Else
            SwitchCondition = IIf(SwitchValue, 0, 4)
        End If
    Else
        SwitchCondition = IIf(SwitchValue, 3, 7)
    End If
    Call Refresh
    RaiseEvent Click(SwitchValue)
End Property
Public Property Get skin() As mySkin
    skin = SwitchSkin
End Property
Public Property Let skin(ByVal vNewValue As mySkin)
    SwitchSkin = vNewValue
    Call ChangeSkin
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("value", True)
    Enable = PropBag.ReadProperty("enable", True)
    SwitchSkin = PropBag.ReadProperty("skin", 0)
    ChangeSkin
    Timer1.Enabled = Ambient.UserMode
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("value", SwitchValue, True)
    Call PropBag.WriteProperty("enable", SwitchEnable, True)
    Call PropBag.WriteProperty("skin", SwitchSkin, 0)
End Sub

Private Sub UserControl_Resize()
    If SwitchSkin = Safe Then
        UserControl.Width = 78 * 15
        UserControl.Height = 23 * 15
    ElseIf SwitchSkin = Kingsoft Then
        UserControl.Width = 74 * 15
        UserControl.Height = 24 * 15
    Else
        UserControl.Width = 36 * 15
        UserControl.Height = 16 * 15
    End If
End Sub


