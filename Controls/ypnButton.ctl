VERSION 5.00
Begin VB.UserControl ypnButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ToolboxBitmap   =   "ypnButton.ctx":0000
   Begin VB.PictureBox ico 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   3120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox ico 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   3000
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox ico 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   2880
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox ico 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   2760
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   3960
      Picture         =   "ypnButton.ctx":0312
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   3960
      Picture         =   "ypnButton.ctx":0B88
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   3960
      Picture         =   "ypnButton.ctx":13FE
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3960
      Picture         =   "ypnButton.ctx":1C74
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3960
      Picture         =   "ypnButton.ctx":24EA
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3960
      Picture         =   "ypnButton.ctx":2D60
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox skinPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3960
      Picture         =   "ypnButton.ctx":35D6
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "ypnButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ypnButton
' Author    : YPN
' Date      : 2018-03-24 00:03
' Purpose   : 普通按钮
'---------------------------------------------------------------------------------------

Option Explicit
Private Const DT_CENTER         As Long = &H1
Private Const DT_VCENTER        As Long = &H4
Private Const DT_SINGLELINE     As Long = &H20
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type
Enum ypnButton_skin
    skin0 = 0
    skin1 = 1
    skin2 = 2
    skin3 = 3
    skin4 = 4
    skin5 = 5
    skin6 = 6
End Enum
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Dim m_Caption As String
Dim m_FocusRect As Boolean
Dim m_State As Long
Dim m_button As Integer
Dim capture As Boolean
Dim ScaleWidth As Single, ScaleHeight As Single
Dim skin As ypnButton_skin
Dim m_float As Boolean
Dim button_ico As Boolean, icoLeft As Long, icoTop As Long, icoWidth As Long, icoHeight As Long
Dim txtRect As RECT
Dim m_check As Boolean '使用复选模式
Dim m_Value As Boolean '按下/普通，复选模式有效

Public Event Click() '声明Click事件

Private Sub setUnSetIco(icoPic As StdPicture)
    Debug.Print "开始 show"
    Debug.Print icoPic
    Debug.Print ico(1).Picture
    If icoPic <> 0 Then  'ico0有图
        If ico(1).Picture = 0 Then
            Set ico(1).Picture = icoPic
            Debug.Print "设了ico1"
        End If
        If ico(2).Picture = 0 Then Set ico(2).Picture = icoPic
        'If ico(3).Picture = 0 Then Set ico(3).Picture = icoPic
    End If
End Sub

Private Sub UserControl_DblClick()
    If m_button = 1 Then Call DoRedraw(2)
    'RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    Call RedrawButton
End Sub

Private Sub UserControl_LostFocus()
    Call RedrawButton
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_button = Button
    If Button = 1 Then
        If m_check Then m_Value = Not m_Value
        Call DoRedraw(2)
    End If
    'If capture Then
    'Call ReleaseCapture
    'capture = False
    'End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 0 And X < ScaleWidth And Y > 0 And Y < ScaleHeight Then
        If Button = 1 Then Exit Sub
        If Not capture Then
            Call SetCapture(UserControl.hwnd)
            capture = True
            Call DoRedraw(1)
        End If
    Else
        'If capture Then
        Call ReleaseCapture
        capture = False
        Call DoRedraw(0)
        'End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If capture Then
        Call ReleaseCapture
        capture = False
        Call DoRedraw(0)
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Resize()
    If UserControl.ScaleHeight < 7 Or UserControl.ScaleWidth < 7 Then Exit Sub
    Dim i As Integer
    ScaleWidth = UserControl.ScaleWidth
    ScaleHeight = UserControl.ScaleHeight
    skinPicture.Width = ScaleWidth * 4
    skinPicture.Height = ScaleHeight
    Call calcPosition
    skinPicture.Cls
    If m_float Then i = 1 Else i = 0
    For i = i To 3
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth, 0, 3, 3, i * 7, 0, 3, 3, vbSrcCopy
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth, 3, 3, ScaleHeight - 6, i * 7, 3, 3, 19, vbSrcCopy
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth, ScaleHeight - 3, 3, 3, i * 7, 22, 3, 3, vbSrcCopy
        
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth + 3, 0, ScaleWidth - 6, 3, i * 7 + 3, 0, 1, 3, vbSrcCopy
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth + 3, 3, ScaleWidth - 6, ScaleHeight - 6, i * 7 + 3, 3, 1, 19, vbSrcCopy
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth + 3, ScaleHeight - 3, ScaleWidth - 6, 3, i * 7 + 3, 22, 1, 3, vbSrcCopy
        
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth + ScaleWidth - 3, 0, 3, 3, i * 7 + 4, 0, 3, 3, vbSrcCopy
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth + ScaleWidth - 3, 3, 3, ScaleHeight - 6, i * 7 + 4, 3, 3, 19, vbSrcCopy
        skinPicture.PaintPicture Picture1(skin).Image, i * ScaleWidth + ScaleWidth - 3, ScaleHeight - 3, 3, 3, i * 7 + 4, 22, 3, 3, vbSrcCopy
    Next
    Call RedrawButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)  '读取属性
    m_Caption = PropBag.ReadProperty("Caption", m_Caption)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", UserControl.Enabled)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", UserControl.BackColor)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", UserControl.ForeColor)
    m_FocusRect = PropBag.ReadProperty("focusRect", True)
    skin = PropBag.ReadProperty("skin", 0)
    m_State = IIf(UserControl.Enabled, 0, 3)
    Font = PropBag.ReadProperty("Font", UserControl.Font)
    m_float = PropBag.ReadProperty("float", False)
    skinPicture.BackColor = PropBag.ReadProperty("backColor")
    button_ico = PropBag.ReadProperty("useIco")
    icoWidth = PropBag.ReadProperty("icoWidth")
    icoHeight = PropBag.ReadProperty("icoHeight")
    Set ico(0).Picture = PropBag.ReadProperty("ico0")
    Set ico(1).Picture = PropBag.ReadProperty("ico1")
    Set ico(2).Picture = PropBag.ReadProperty("ico2")
    Set ico(3).Picture = PropBag.ReadProperty("ico3")
    m_check = PropBag.ReadProperty("m_check", False)
    m_Value = PropBag.ReadProperty("m_value", False)
    Call DoRedraw(IIf(m_Value, 2, 0))
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)  '写入属性
    Call PropBag.WriteProperty("Caption", m_Caption)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor)
    Call PropBag.WriteProperty("focusRect", m_FocusRect)
    Call PropBag.WriteProperty("skin", skin)
    Call PropBag.WriteProperty("Font", UserControl.Font)
    
    Call PropBag.WriteProperty("float", m_float)
    Call PropBag.WriteProperty("backColor", skinPicture.BackColor)
    Call PropBag.WriteProperty("useIco", button_ico)
    Call PropBag.WriteProperty("icoWidth", icoWidth)
    Call PropBag.WriteProperty("icoHeight", icoHeight)
    Call PropBag.WriteProperty("ico0", ico(0).Picture)
    Call PropBag.WriteProperty("ico1", ico(1).Picture)
    Call PropBag.WriteProperty("ico2", ico(2).Picture)
    Call PropBag.WriteProperty("ico3", ico(3).Picture)
    Call PropBag.WriteProperty("m_check", m_check)
    Call PropBag.WriteProperty("m_value", m_Value)
End Sub

Private Sub DoRedraw(ByVal nState As Long)   '0普通  1高亮 2按下 3无效
    If m_State = nState Then Exit Sub
    m_State = nState
    If m_check And m_State <> 3 Then
        If m_Value Then m_State = 2  'true 按下
    End If
    Call RedrawButton
End Sub

Private Sub RedrawButton()
    UserControl.Cls     '先清除用户控件上的旧内容
    UserControl.PaintPicture skinPicture.Image, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_State * UserControl.ScaleWidth, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, vbSrcCopy
    If UserControl.Enabled Then
        UserControl.ForeColor = GetSysColor(9)
    Else
        UserControl.ForeColor = GetSysColor(19)
    End If
    Call calcPosition
    
    If button_ico Then DrawIconEx UserControl.hDC, icoLeft, icoTop, ico(m_State).Picture, icoWidth, icoHeight, 0, 0, DI_NORMAL
    Call SetTextColor(UserControl.hDC, &H0) '让FocusRect是黑色的。。
    Call DrawText(UserControl.hDC, m_Caption, -1, txtRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    
    If GetFocus = UserControl.hwnd And m_FocusRect Then    '如果有焦点，则绘制FocusRect
        Call SetTextColor(UserControl.hDC, &H0) '让FocusRect是黑色的。。
        Call SetRect(txtRect, 3, 3, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3)
        Call DrawFocusRect(UserControl.hDC, txtRect)
    End If
    UserControl.Refresh
End Sub

Private Sub UserControl_InitProperties()   '初始化控件大小
    m_FocusRect = False
    m_Caption = Ambient.DisplayName
    skin = skin0
    m_float = False
    skinPicture.BackColor = &H80000005
    button_ico = False
    icoLeft = 0
    icoTop = 0
    
    m_check = False
    m_Value = False
End Sub

Public Property Get Caption() As String    '返回标题
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal newValue As String) '设置标题
    m_Caption = Trim$(newValue)
    Call RedrawButton
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean) '设置可用状态
    UserControl.Enabled = newValue
    m_State = IIf(newValue, 0, 3)
    Call RedrawButton
End Property

Public Property Get Font() As StdFont  '返回字体
    Set Font = UserControl.Font
End Property

Public Property Let Font(ByVal newValue As StdFont)  '设置字体
    Set UserControl.Font = newValue
    Call RedrawButton
End Property

Public Property Set Font(ByVal newValue As StdFont)  '设置字体
    Set UserControl.Font = newValue
    Call RedrawButton
End Property

Public Property Get BackColor() As OLE_COLOR  '返回背景色
    BackColor = skinPicture.BackColor
End Property

Public Property Let BackColor(ByVal newValue As OLE_COLOR)    '设置背景色
    'UserControl.BackColor = NewValue
    skinPicture.BackColor = newValue
    Call UserControl_Resize
End Property

Public Property Get ForeColor() As OLE_COLOR '返回前景色
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal newValue As OLE_COLOR)   '设置前景色
    UserControl.ForeColor = newValue
    Call RedrawButton
End Property

Public Property Get FocusRect() As Boolean
    FocusRect = m_FocusRect
End Property

Public Property Let FocusRect(ByVal newValue As Boolean)
    m_FocusRect = newValue
End Property

Public Property Get checkButton() As Boolean
    checkButton = m_check
End Property

Public Property Let checkButton(ByVal newValue As Boolean)
    m_check = newValue
    Call RedrawButton
End Property

Public Property Get checkValue() As Boolean
    checkValue = m_Value
End Property

Public Property Let checkValue(ByVal newValue As Boolean)
    If m_check Then
        m_Value = newValue
        Call DoRedraw(IIf(m_Value, 2, 0))
    End If
End Property

Public Property Get buttonSkin() As ypnButton_skin
Attribute buttonSkin.VB_Description = "皮肤"
    buttonSkin = skin
End Property

Public Property Let buttonSkin(ByVal newValue As ypnButton_skin)
    If skin <> newValue Then
        skin = newValue
        Call UserControl_Resize
    End If
End Property

Public Property Get useIco() As Boolean
    useIco = button_ico
End Property

Public Property Let useIco(ByVal newValue As Boolean)
    If button_ico <> newValue Then
        button_ico = newValue
        Call UserControl_Resize
    End If
End Property

Public Property Get Float() As Boolean
    Float = m_float
End Property

Public Property Let Float(ByVal newValue As Boolean)
    If m_float <> newValue Then
        m_float = newValue
        Call UserControl_Resize
    End If
End Property

Public Property Get ico0() As StdPicture
Attribute ico0.VB_Description = "普通状态显示的图标"
    Set ico0 = ico(0).Picture
End Property

Public Property Set ico0(ByVal newValue As StdPicture)
    Set ico(0).Picture = newValue
    icoWidth = ico(0).ScaleWidth
    icoHeight = ico(0).ScaleHeight
    button_ico = True
    setUnSetIco newValue
    Call UserControl_Resize
End Property

Public Property Get ico1() As StdPicture
Attribute ico1.VB_Description = "鼠标指向时显示的图标"
    Set ico1 = ico(1).Picture
End Property

Public Property Set ico1(ByVal newValue As StdPicture)
    Set ico(1).Picture = newValue
    icoWidth = ico(1).ScaleWidth
    icoHeight = ico(1).ScaleHeight
    button_ico = True
    Call UserControl_Resize
End Property

Public Property Get ico2() As StdPicture
Attribute ico2.VB_Description = "按下时显示的图标"
    Set ico2 = ico(2).Picture
End Property

Public Property Set ico2(ByVal newValue As StdPicture)
    Set ico(2).Picture = newValue
    icoWidth = ico(2).ScaleWidth
    icoHeight = ico(2).ScaleHeight
    button_ico = True
    Call UserControl_Resize
End Property

Public Property Get ico3() As StdPicture
Attribute ico3.VB_Description = "无效时显示的图标"
    Set ico3 = ico(3).Picture
End Property

Public Property Set ico3(ByVal newValue As StdPicture)
    Set ico(3).Picture = newValue
    icoWidth = ico(3).ScaleWidth
    icoHeight = ico(3).ScaleHeight
    button_ico = True
    Call UserControl_Resize
End Property

Private Sub calcPosition()
    If button_ico Then
        If Len(m_Caption) > 0 Then
            Dim FontHeight As Long
            FontHeight = UserControl.FontSize
            If m_State = 2 Then '按下
                icoLeft = (UserControl.ScaleWidth - icoWidth) \ 2 + 2
                icoTop = (UserControl.ScaleHeight - icoHeight) \ 2 - FontHeight + 2
                If icoTop < 3 Then icoTop = 4
                Call SetRect(txtRect, 5, icoTop + icoHeight + 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)
            Else
                icoLeft = (UserControl.ScaleWidth - icoWidth) \ 2
                icoTop = (UserControl.ScaleHeight - icoHeight) \ 2 - FontHeight
                If icoTop < 1 Then icoTop = 2
                Call SetRect(txtRect, 3, icoTop + icoHeight + 0, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2)
            End If
        Else
            icoLeft = (UserControl.ScaleWidth - icoWidth) \ 2
            icoTop = (UserControl.ScaleHeight - icoHeight) \ 2
            If m_State = 2 Then
                icoLeft = icoLeft + 1
                icoTop = icoTop + 1
            End If
        End If
    Else
        If m_State = 2 Then
            Call SetRect(txtRect, 5, 5, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3)
        Else
            Call SetRect(txtRect, 3, 3, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3)
        End If
    End If
End Sub
