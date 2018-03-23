VERSION 5.00
Begin VB.UserControl ypnCheckBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   8955
   ToolboxBitmap   =   "ypnCheckBox.ctx":0000
   Begin VB.Image y16 
      Height          =   240
      Left            =   5775
      Picture         =   "ypnCheckBox.ctx":0312
      Stretch         =   -1  'True
      Top             =   6615
      Width           =   240
   End
   Begin VB.Image g16 
      Height          =   240
      Left            =   5250
      Picture         =   "ypnCheckBox.ctx":0368
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   240
   End
   Begin VB.Image b16 
      Height          =   240
      Left            =   4725
      Picture         =   "ypnCheckBox.ctx":03BF
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   240
   End
   Begin VB.Image r16 
      Height          =   240
      Left            =   3990
      Picture         =   "ypnCheckBox.ctx":0415
      Stretch         =   -1  'True
      Top             =   6825
      Width           =   240
   End
   Begin VB.Image y8 
      Height          =   120
      Left            =   7770
      Picture         =   "ypnCheckBox.ctx":046B
      Stretch         =   -1  'True
      Top             =   5985
      Width           =   120
   End
   Begin VB.Image b8 
      Height          =   120
      Left            =   7245
      Picture         =   "ypnCheckBox.ctx":04AD
      Stretch         =   -1  'True
      Top             =   5985
      Width           =   120
   End
   Begin VB.Image g8 
      Height          =   120
      Left            =   7770
      Picture         =   "ypnCheckBox.ctx":04EF
      Stretch         =   -1  'True
      Top             =   5670
      Width           =   120
   End
   Begin VB.Image r8 
      Height          =   120
      Left            =   7245
      Picture         =   "ypnCheckBox.ctx":0531
      Stretch         =   -1  'True
      Top             =   5670
      Width           =   120
   End
   Begin VB.Image bl128 
      Height          =   1920
      Left            =   2640
      Picture         =   "ypnCheckBox.ctx":0573
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image y128 
      Height          =   1920
      Left            =   120
      Picture         =   "ypnCheckBox.ctx":51BD
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image g128 
      Height          =   1920
      Left            =   2760
      Picture         =   "ypnCheckBox.ctx":9E07
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image b128 
      Height          =   1920
      Left            =   5640
      Picture         =   "ypnCheckBox.ctx":EA51
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image r128 
      Height          =   1920
      Left            =   5565
      Picture         =   "ypnCheckBox.ctx":1369B
      Stretch         =   -1  'True
      Top             =   3150
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Shape Border 
      BackColor       =   &H00C0E0FF&
      Height          =   1935
      Left            =   10
      Shape           =   5  'Rounded Square
      Top             =   10
      Width           =   1815
   End
   Begin VB.Shape Light 
      BackColor       =   &H0060FFFF&
      BorderColor     =   &H0060FFFF&
      Height          =   1215
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image imgPic 
      Height          =   1275
      Left            =   5355
      Stretch         =   -1  'True
      Top             =   5250
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Image imgMouseOn 
      Height          =   1275
      Left            =   3465
      Stretch         =   -1  'True
      Top             =   4935
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Image imgMouseDown 
      Height          =   1275
      Left            =   1890
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Image imgInvalid 
      Height          =   1275
      Left            =   315
      Stretch         =   -1  'True
      Top             =   5040
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Shape Shadow 
      BorderColor     =   &H00404040&
      Height          =   1095
      Left            =   5880
      Shape           =   5  'Rounded Square
      Top             =   2100
      Width           =   1815
   End
End
Attribute VB_Name = "ypnCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Dim m_TransParent
Dim Min As Boolean
Dim mDown As Boolean
'缺省属性值:
Const m_def_usePicture = False
Const m_def_style = 1
Const m_def_ShowBorder = True
'Const m_def_style = 0
Const m_def_TransParent = True
'Const m_def_borderLight = True
Const m_def_borderShadow = True
Const m_def_Value = False
Const m_def_OLEDragMode = 0
'属性变量:
Dim m_usePicture As Boolean
Dim m_style As Long
Dim m_ShowBorder As Boolean
'Dim m_Picture As Picture
'Dim m_PicMouseOn As Picture
'Dim m_PicMouseDown As Picture
'Dim m_PicInvalid As Picture
'Dim m_style As Long
Dim m_TransParent As Boolean
'Dim m_borderLight As Boolean
Dim m_borderShadow As Boolean
Dim m_Value As Boolean
Dim m_OLEDragMode As Integer
'事件声明:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "当用户在一个对象上按下并释放鼠标按钮时发生。"
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "当第一次显示一个窗体时或改变一个对象的大小时发生。"
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "当用户在拥有焦点的对象上按下鼠标按钮时发生。"
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "当用户移动鼠标时发生。"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "当用户在拥有焦点的对象上释放鼠标发生。"
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "当完成或取消手动或自动拖/放之后，在 OLE 拖/放源控件上发生。"
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "OLEDropMode 的属性设置为手动、且数据通过 OLE 拖/放操作放入控件时发生。"
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "当 OLEDropMode 属性设置为手动、且 OLE 拖/放操作期间鼠标经过控件时发生。"
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "当需要改变鼠标光标时，在 OLE 拖/放操作中的源控件上发生。"
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "在 OLEDragStart 事件期间，放下目标所需的数据未提供给 DataObject 时，在 OLE 拖/放源控件上发生。"
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "手动或自动初始化 OLE 拖/放操作时发生。"
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "当移动、放大或露出图片框的任何部分时发生。"

Public Event MouseIn()
Public Event MouseExit()

Enum cShape
    Rectangle = 0
    square = 1
    oval = 2
    shapeCircle = 3
    roundedrect = 4
    roundedsquare = 5
End Enum


Enum cStyle
    Red = 0
    Green = 1
    Blue = 2
    yellow = 3
    black = 4
    None = 5
End Enum


'注意！不要删除或修改下列被注释的行！
'MappingInfo=Border,Border,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = Border.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Border.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    ReDraw
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Border,Border,-1,BorderColor
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "返回/设置对象的边框颜色。"
    BorderColor = Border.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
    Border.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
    ReDraw
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Border,Border,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "返回/设置对象的边框样式。"
    BorderStyle = Border.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    Border.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    ReDraw
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Border,Border,-1,BorderWidth
Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "返回/设置控件的边框宽度。"
    BorderWidth = Border.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    Border.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
    ReDraw
End Property



Private Sub b128_Click()
    UserControl_Click
End Sub

Private Sub b16_Click()
    UserControl_Click
End Sub

Private Sub b8_Click()
    UserControl_Click
End Sub

Private Sub bl128_Click()
    UserControl_Click
End Sub

Private Sub g128_Click()
    UserControl_Click
End Sub

Private Sub g16_Click()
    UserControl_Click
End Sub

Private Sub g8_Click()
    UserControl_Click
End Sub

Private Sub r128_Click()
    UserControl_Click
End Sub

Private Sub r16_Click()
    UserControl_Click
End Sub

Private Sub r8_Click()
    UserControl_Click
End Sub

Private Sub UserControl_Click()
    m_Value = Not m_Value
    Debug.Print "value" & m_Value
    ReDraw
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    ReDraw
End Property

Private Sub UserControl_Resize()
    
    
    If UserControl.Height <= 140 Or UserControl.Width <= 140 Then
        UserControl.Height = 140
        UserControl.Width = 140
    End If
    If UserControl.Height > 140 Then
        If UserControl.Height < 260 Or UserControl.Width < 260 Then
            UserControl.Height = 260
            UserControl.Width = 260
        End If
    End If
    
    
    ReDraw
    RaiseEvent Resize
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mDown = True
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If (X < 0) Or (Y < 0) Or (X > UserControl.Width) Or (Y > UserControl.Height) Then
        ReleaseCapture
        Min = False
        ReDraw
        RaiseEvent MouseExit  ' 鼠标离开的代码
    Else
        SetCapture hwnd
        Min = True
        ReDraw
        RaiseEvent MouseIn   ' 鼠标进入的代码
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mDown = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "设置一个自定义鼠标图标。"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "以给定控件作为源，启动一个 OLE 拖/放事件。"
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get OLEDragMode() As Integer
Attribute OLEDragMode.VB_Description = "返回/设置该对象是否能作为 OLE 拖/放源，以及该进程是自动启动，还是在程序控制下启动。"
    OLEDragMode = m_OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As Integer)
    m_OLEDragMode = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "返回/设置该对象是否能作为一个 OLE 放下目标。"
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Border,Border,-1,Shape
Public Property Get Shape() As cShape
Attribute Shape.VB_Description = "返回/设置一个值，指出控件外观。"
    Shape = Border.Shape
End Property

Public Property Let Shape(ByVal New_Shape As cShape)
    Light.Shape = New_Shape
    Shadow.Shape = New_Shape
    Border.Shape() = New_Shape
    PropertyChanged "Shape"
    ReDraw
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_OLEDragMode = m_def_OLEDragMode
    m_TransParent = m_def_TransParent
    '     m_borderLight = m_def_borderLight
    m_borderShadow = m_def_borderShadow
    m_Value = m_def_Value
    '     m_style = m_def_style
    '     Set m_Picture = LoadPicture("")
    '     Set m_PicMouseOn = LoadPicture("")
    '     Set m_PicMouseDown = LoadPicture("")
    '     Set m_PicInvalid = LoadPicture("")
    m_style = m_def_style
    m_ShowBorder = m_def_ShowBorder
    m_usePicture = m_def_usePicture
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Border.BackColor = PropBag.ReadProperty("BackColor", &H80FFFF)
    Border.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    Border.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Border.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    '     m_TransParent = PropBag.ReadProperty("Transparent", False)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_OLEDragMode = PropBag.ReadProperty("OLEDragMode", m_def_OLEDragMode)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Border.Shape = PropBag.ReadProperty("Shape", 5)
    m_TransParent = PropBag.ReadProperty("TransParent", m_def_TransParent)
    '     m_borderLight = PropBag.ReadProperty("borderLight", m_def_borderLight)
    m_borderShadow = PropBag.ReadProperty("borderShadow", m_def_borderShadow)
    m_Value = PropBag.ReadProperty("value", m_def_Value)
    '     m_style = PropBag.ReadProperty("style", m_def_style)
    Light.BorderColor = PropBag.ReadProperty("lightColor", 6356991)
    Shadow.BorderColor = PropBag.ReadProperty("shadowColor", 4210752)
    '     Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    '     Set m_PicMouseOn = PropBag.ReadProperty("PicMouseOn", Nothing)
    '     Set m_PicMouseDown = PropBag.ReadProperty("PicMouseDown", Nothing)
    '     Set m_PicInvalid = PropBag.ReadProperty("PicInvalid", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set Picture = PropBag.ReadProperty("PicMouseOn", Nothing)
    Set Picture = PropBag.ReadProperty("PicMouseDown", Nothing)
    Set Picture = PropBag.ReadProperty("PicInvalid", Nothing)
    m_style = PropBag.ReadProperty("style", m_def_style)
    m_ShowBorder = PropBag.ReadProperty("ShowBorder", m_def_ShowBorder)
    m_usePicture = PropBag.ReadProperty("usePicture", m_def_usePicture)
End Sub

Private Sub UserControl_Show()
    ReDraw
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("BackColor", Border.BackColor, &H80FFFF)
    Call PropBag.WriteProperty("BorderColor", Border.BorderColor, -2147483640)
    Call PropBag.WriteProperty("BorderStyle", Border.BorderStyle, 1)
    Call PropBag.WriteProperty("BorderWidth", Border.BorderWidth, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("OLEDragMode", m_OLEDragMode, m_def_OLEDragMode)
    '      Call PropBag.WriteProperty("Transparent", m_TransParent, False)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("Shape", Border.Shape, 5)
    Call PropBag.WriteProperty("TransParent", m_TransParent, m_def_TransParent)
    '     Call PropBag.WriteProperty("borderLight", m_borderLight, m_def_borderLight)
    Call PropBag.WriteProperty("borderShadow", m_borderShadow, m_def_borderShadow)
    Call PropBag.WriteProperty("value", m_Value, m_def_Value)
    '     Call PropBag.WriteProperty("style", m_style, m_def_style)
    Call PropBag.WriteProperty("lightColor", Light.BorderColor, 6356991)
    Call PropBag.WriteProperty("shadowColor", Shadow.BorderColor, 4210752)
    '     Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    '     Call PropBag.WriteProperty("PicMouseOn", m_PicMouseOn, Nothing)
    '     Call PropBag.WriteProperty("PicMouseDown", m_PicMouseDown, Nothing)
    '     Call PropBag.WriteProperty("PicInvalid", m_PicInvalid, Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("PicMouseOn", Picture, Nothing)
    Call PropBag.WriteProperty("PicMouseDown", Picture, Nothing)
    Call PropBag.WriteProperty("PicInvalid", Picture, Nothing)
    Call PropBag.WriteProperty("style", m_style, m_def_style)
    Call PropBag.WriteProperty("ShowBorder", m_ShowBorder, m_def_ShowBorder)
    Call PropBag.WriteProperty("usePicture", m_usePicture, m_def_usePicture)
End Sub

Private Sub ReDraw()
    UserControl.Cls
    
    If m_TransParent Then
        Border.BackStyle = 0
        CopyBackGround UserControl.Parent.hwnd, UserControl.hwnd, UserControl.hDC '复制父窗背景透明
        Debug.Print "透明"
    Else
        
        If m_usePicture Then
            Border.BackStyle = 0
            
            If Enabled Then
                imgPic.Height = UserControl.Height
                imgPic.Width = UserControl.Width
                imgMouseDown.Height = UserControl.Height
                imgMouseDown.Width = UserControl.Width
                imgMouseOn.Height = UserControl.Height
                imgMouseOn.Width = UserControl.Width
                imgInvalid.Height = UserControl.Height
                imgInvalid.Width = UserControl.Width
                
                If Min And mDown = False Then
                    UserControl.PaintPicture imgMouseOn, 0, 0, UserControl.Width, UserControl.Height
                ElseIf mDown Then
                    UserControl.PaintPicture imgMouseDown, 0, 0, UserControl.Width, UserControl.Height
                Else
                    UserControl.PaintPicture imgPic, 0, 0, UserControl.Width, UserControl.Height
                    
                End If
                
            Else
                UserControl.PaintPicture imgInvalid, 0, 0, UserControl.Width, UserControl.Height
                
            End If
            
        Else
            Border.BackStyle = 1
            
        End If
        
    End If
    
    UserControl.Refresh
    
    Light.Height = UserControl.Height
    Light.Width = UserControl.Width
    Border.Width = UserControl.Width - 10
    Border.Height = UserControl.Height - 10
    Shadow.Left = Border.Left + Border.BorderWidth * 9.5
    Shadow.Top = Border.Top + Border.BorderWidth * 9.5
    Shadow.Height = UserControl.Height - Border.BorderWidth * 9.5 - 10
    Shadow.Width = UserControl.Width - Border.BorderWidth * 9.5 - 10
    Light.Shape = Border.Shape
    Shadow.Shape = Border.Shape
    
    If Min Then
        Light.Visible = True
    Else
        Light.Visible = False
        
    End If
    
    'Light.Visible = m_borderLight
    Shadow.Visible = m_borderShadow
    Border.Visible = m_ShowBorder
    
    If m_Value Then
        
        Select Case m_style
            
        Case 0  '红勾
            
            If UserControl.Height <= 260 And UserControl.Height > 140 Then
                r16.Left = Border.Left + Border.BorderWidth * 9 - 10
                r16.Top = Border.Top + Border.BorderWidth * 7
                r16.Width = Border.Width - 10 - Border.BorderWidth * 10
                r16.Height = Border.Height - 10 - Border.BorderWidth * 10
                r16.Visible = True
                r16.ZOrder
                r8.Visible = False
                r128.Visible = False
            ElseIf UserControl.Height <= 140 Then
                r8.Left = Border.Left + Border.BorderWidth * 9 - 10
                r8.Top = Border.Top + Border.BorderWidth * 5
                r8.Width = Border.Width - 10 - Border.BorderWidth * 10
                r8.Height = Border.Height - 10 - Border.BorderWidth * 10
                r8.Visible = True
                r8.ZOrder
                r128.Visible = False
                r16.Visible = False
            Else
                r128.Left = Border.Left + Border.BorderWidth * 9.5 - 10
                r128.Top = Border.Top + Border.BorderWidth * 9.5 + 10
                r128.Width = Border.Width - 10 - Border.BorderWidth * 10
                r128.Height = Border.Height - 10 - Border.BorderWidth * 10
                r128.Visible = True
                r128.ZOrder
                r8.Visible = False
                r16.Visible = False
                
            End If
            
            g8.Visible = False
            b8.Visible = False
            y8.Visible = False
            g16.Visible = False
            b16.Visible = False
            y16.Visible = False
            g128.Visible = False
            b128.Visible = False
            y128.Visible = False
            bl128.Visible = False
            'UserControl.PaintPicture r128, 0, 0, UserControl.Width, UserControl.Height
            
        Case 1
            
            If UserControl.Height <= 260 And UserControl.Height > 140 Then
                g16.Left = Border.Left + Border.BorderWidth * 9 - 10
                g16.Top = Border.Top + Border.BorderWidth * 7
                g16.Width = Border.Width - 10 - Border.BorderWidth * 10
                g16.Height = Border.Height - 10 - Border.BorderWidth * 10
                g16.Visible = True
                g16.ZOrder
                g8.Visible = False
                g128.Visible = False
            ElseIf UserControl.Height <= 140 Then
                g8.Left = Border.Left + Border.BorderWidth * 9 - 10
                g8.Top = Border.Top + Border.BorderWidth * 5
                g8.Width = Border.Width - 10 - Border.BorderWidth * 10
                g8.Height = Border.Height - 10 - Border.BorderWidth * 10
                g8.Visible = True
                g8.ZOrder
                g128.Visible = False
                g16.Visible = False
            Else
                g128.Left = Border.Left + Border.BorderWidth * 9.5 - 10
                g128.Top = Border.Top + Border.BorderWidth * 9.5 + 10
                g128.Width = Border.Width - 10 - Border.BorderWidth * 10
                g128.Height = Border.Height - 10 - Border.BorderWidth * 10
                g128.Visible = True
                g128.ZOrder
                g8.Visible = False
                g16.Visible = False
                
            End If
            
            r8.Visible = False
            b8.Visible = False
            y8.Visible = False
            r16.Visible = False
            b16.Visible = False
            y16.Visible = False
            r128.Visible = False
            b128.Visible = False
            y128.Visible = False
            bl128.Visible = False
            'UserControl.PaintPicture g128, 0, 0, UserControl.Width, UserControl.Height
            
        Case 2
            
            If UserControl.Height <= 260 And UserControl.Height > 140 Then
                b16.Left = Border.Left + Border.BorderWidth * 9 - 10
                b16.Top = Border.Top + Border.BorderWidth * 7
                b16.Width = Border.Width - 10 - Border.BorderWidth * 10
                b16.Height = Border.Height - 10 - Border.BorderWidth * 10
                b16.Visible = True
                b16.ZOrder
                b8.Visible = False
                b128.Visible = False
            ElseIf UserControl.Height <= 140 Then
                b8.Left = Border.Left + Border.BorderWidth * 9 - 10
                b8.Top = Border.Top + Border.BorderWidth * 5
                b8.Width = Border.Width - 10 - Border.BorderWidth * 10
                b8.Height = Border.Height - 10 - Border.BorderWidth * 10
                b8.Visible = True
                b8.ZOrder
                b128.Visible = False
                b16.Visible = False
            Else
                b128.Left = Border.Left + Border.BorderWidth * 9.5 - 10
                b128.Top = Border.Top + Border.BorderWidth * 9.5 + 10
                b128.Width = Border.Width - 10 - Border.BorderWidth * 10
                b128.Height = Border.Height - 10 - Border.BorderWidth * 10
                b128.Visible = True
                b128.ZOrder
                b8.Visible = False
                b16.Visible = False
                
            End If
            
            g8.Visible = False
            r8.Visible = False
            y8.Visible = False
            g16.Visible = False
            r16.Visible = False
            y16.Visible = False
            g128.Visible = False
            r128.Visible = False
            y128.Visible = False
            bl128.Visible = False
            'UserControl.PaintPicture b128, 0, 0, UserControl.Width, UserControl.Height
            
        Case 3  '黑勾
            bl128.Left = Border.Left + Border.BorderWidth * 9.5 - 10
            bl128.Top = Border.Top + Border.BorderWidth * 9.5 + 10
            bl128.Width = Border.Width - 10 - Border.BorderWidth * 10
            bl128.Height = Border.Height - 10 - Border.BorderWidth * 10
            bl128.Visible = True
            bl128.ZOrder
            b8.Visible = False
            g8.Visible = False
            r8.Visible = False
            y8.Visible = False
            b16.Visible = False
            g16.Visible = False
            r16.Visible = False
            y16.Visible = False
            b128.Visible = False
            g128.Visible = False
            r128.Visible = False
            y128.Visible = False
            
            'UserControl.PaintPicture bl128, 0, 0, UserControl.Width, UserControl.Height
            
        Case 4
            
            If UserControl.Height <= 260 And UserControl.Height > 140 Then
                y16.Left = Border.Left + Border.BorderWidth * 9 - 10
                y16.Top = Border.Top + Border.BorderWidth * 7
                y16.Width = Border.Width - 10 - Border.BorderWidth * 10
                y16.Height = Border.Height - 10 - Border.BorderWidth * 10
                y16.Visible = True
                y16.ZOrder
                y8.Visible = False
                y128.Visible = False
            ElseIf UserControl.Height <= 140 Then
                y8.Left = Border.Left + Border.BorderWidth * 9 - 10
                y8.Top = Border.Top + Border.BorderWidth * 5
                y8.Width = Border.Width - 10 - Border.BorderWidth * 10
                y8.Height = Border.Height - 10 - Border.BorderWidth * 10
                y8.Visible = True
                y8.ZOrder
                y128.Visible = False
                y16.Visible = False
            Else
                y128.Left = Border.Left + Border.BorderWidth * 9.5 - 10
                y128.Top = Border.Top + Border.BorderWidth * 9.5 + 10
                y128.Width = Border.Width - 10 - Border.BorderWidth * 10
                y128.Height = Border.Height - 10 - Border.BorderWidth * 10
                y128.Visible = True
                y128.ZOrder
                y8.Visible = False
                y16.Visible = False
                
            End If
            
            g8.Visible = False
            r8.Visible = False
            b8.Visible = False
            g16.Visible = False
            r16.Visible = False
            b16.Visible = False
            g128.Visible = False
            r128.Visible = False
            b128.Visible = False
            bl128.Visible = False
            'UserControl.PaintPicture y128, 0, 0, UserControl.Width, UserControl.Height
            
        Case 5
            b8.Visible = False
            g8.Visible = False
            r8.Visible = False
            y8.Visible = False
            b16.Visible = False
            g16.Visible = False
            r16.Visible = False
            y16.Visible = False
            b128.Visible = False
            g128.Visible = False
            r128.Visible = False
            y128.Visible = False
            bl128.Visible = False
            
        End Select
        
    Else
        b8.Visible = False
        g8.Visible = False
        r8.Visible = False
        y8.Visible = False
        b16.Visible = False
        g16.Visible = False
        r16.Visible = False
        y16.Visible = False
        b128.Visible = False
        g128.Visible = False
        r128.Visible = False
        y128.Visible = False
        bl128.Visible = False
        
    End If
    
End Sub

'
'
'Public Property Get TransParent() As Boolean
' TransParent = m_TransParent
'End Property
'
'Public Property Let TransParent(ByVal vNewValue As Boolean)
'm_TransParent = vNewValue
'End Property
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,true
Public Property Get TransParent() As Boolean
    TransParent = m_TransParent
End Property

Public Property Let TransParent(ByVal New_TransParent As Boolean)
    m_TransParent = New_TransParent
    PropertyChanged "TransParent"
    ReDraw
End Property
'
''注意！不要删除或修改下列被注释的行！
''MemberInfo=0,0,0,true
'Public Property Get borderLight() As Boolean
'     borderLight = m_borderLight
'End Property
'
'Public Property Let borderLight(ByVal New_borderLight As Boolean)
'     m_borderLight = New_borderLight
'     PropertyChanged "borderLight"
'End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,true
Public Property Get borderShadow() As Boolean
    borderShadow = m_borderShadow
End Property

Public Property Let borderShadow(ByVal New_borderShadow As Boolean)
    m_borderShadow = New_borderShadow
    PropertyChanged "borderShadow"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,false
Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "value"
    ReDraw
End Property
'
''注意！不要删除或修改下列被注释的行！
''MemberInfo=8,0,0,0
'Public Property Get style() As Long
'     style = m_style
'End Property
'
'Public Property Let style(ByVal New_style As cStyle)
'     m_style = New_style
'     PropertyChanged "style"
'End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Light,Light,-1,BorderColor
Public Property Get lightColor() As Long
Attribute lightColor.VB_Description = "返回/设置对象的边框颜色。"
    lightColor = Light.BorderColor
End Property

Public Property Let lightColor(ByVal New_lightColor As Long)
    Light.BorderColor() = New_lightColor
    PropertyChanged "lightColor"
    ReDraw
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Shadow,Shadow,-1,BorderColor
Public Property Get shadowColor() As Long
Attribute shadowColor.VB_Description = "返回/设置对象的边框颜色。"
    shadowColor = Shadow.BorderColor
End Property

Public Property Let shadowColor(ByVal New_shadowColor As Long)
    Shadow.BorderColor() = New_shadowColor
    PropertyChanged "shadowColor"
    ReDraw
End Property
''
'''注意！不要删除或修改下列被注释的行！
'''MemberInfo=11,0,0,0
''Public Property Get Picture() As Picture
''     Set Picture = m_Picture
''End Property
''
''Public Property Set Picture(ByVal New_Picture As Picture)
''     Set m_Picture = New_Picture
''     PropertyChanged "Picture"
''End Property
''
'''注意！不要删除或修改下列被注释的行！
'''MemberInfo=11,0,0,0
''Public Property Get PicMouseOn() As Picture
''     Set PicMouseOn = m_PicMouseOn
''End Property
''
''Public Property Set PicMouseOn(ByVal New_PicMouseOn As Picture)
''     Set m_PicMouseOn = New_PicMouseOn
''     PropertyChanged "PicMouseOn"
''End Property
''
'''注意！不要删除或修改下列被注释的行！
'''MemberInfo=11,0,0,0
''Public Property Get PicMouseDown() As Picture
''     Set PicMouseDown = m_PicMouseDown
''End Property
''
''Public Property Set PicMouseDown(ByVal New_PicMouseDown As Picture)
''     Set m_PicMouseDown = New_PicMouseDown
''     PropertyChanged "PicMouseDown"
''End Property
''
'''注意！不要删除或修改下列被注释的行！
'''MemberInfo=11,0,0,0
''Public Property Get PicInvalid() As Picture
''     Set PicInvalid = m_PicInvalid
''End Property
''
''Public Property Set PicInvalid(ByVal New_PicInvalid As Picture)
''     Set m_PicInvalid = New_PicInvalid
''     PropertyChanged "PicInvalid"
''End Property
''
''注意！不要删除或修改下列被注释的行！
''MappingInfo=imgpic,imgpic,-1,Picture
'Public Property Get Picture() As Picture
'     Set Picture = imgPic.Picture
'End Property
'
'Public Property Set Picture(ByVal New_Picture As Picture)
'     Set imgPic.Picture = New_Picture
'     PropertyChanged "Picture"
'End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=imgMouseOn,imgMouseOn,-1,Picture
Public Property Get PicMouseOn() As Picture
Attribute PicMouseOn.VB_Description = "返回/设置控件中显示的图形。"
    Set PicMouseOn = imgMouseOn.Picture
End Property

Public Property Set PicMouseOn(ByVal New_PicMouseOn As Picture)
    Set imgMouseOn.Picture = New_PicMouseOn
    PropertyChanged "PicMouseOn"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=imgMouseDown,imgMouseDown,-1,Picture
Public Property Get PicMouseDown() As Picture
Attribute PicMouseDown.VB_Description = "返回/设置控件中显示的图形。"
    Set PicMouseDown = imgMouseDown.Picture
End Property

Public Property Set PicMouseDown(ByVal New_PicMouseDown As Picture)
    Set imgMouseDown.Picture = New_PicMouseDown
    PropertyChanged "PicMouseDown"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=imgInvalid,imgInvalid,-1,Picture
Public Property Get PicInvalid() As Picture
Attribute PicInvalid.VB_Description = "返回/设置控件中显示的图形。"
    Set PicInvalid = imgInvalid.Picture
End Property

Public Property Set PicInvalid(ByVal New_PicInvalid As Picture)
    Set imgInvalid.Picture = New_PicInvalid
    PropertyChanged "PicInvalid"
    ReDraw
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=imgPic,imgPic,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "返回/设置控件中显示的图形。"
    Set Picture = imgPic.Picture
    
End Property

Public Property Set Picture(ByVal New_Pic As Picture)
    
    If imgPic.Picture > 0 Then
        m_usePicture = True
    Else
        m_usePicture = False
    End If
    ReDraw
    'Set imgPic.Picture = New_Pic
    'PropertyChanged "Picture"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,1
Public Property Get Style() As cStyle
    Style = m_style
End Property

Public Property Let Style(ByVal New_Style As cStyle)
    m_style = New_Style
    PropertyChanged "style"
    ReDraw
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,true
Public Property Get ShowBorder() As Boolean
    ShowBorder = m_ShowBorder
End Property

Public Property Let ShowBorder(ByVal New_ShowBorder As Boolean)
    m_ShowBorder = New_ShowBorder
    PropertyChanged "ShowBorder"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,false
Public Property Get usePicture() As Boolean
    usePicture = m_usePicture
End Property

Public Property Let usePicture(ByVal New_usePicture As Boolean)
    m_usePicture = New_usePicture
    PropertyChanged "usePicture"
End Property

Private Sub y128_Click()
    UserControl_Click
End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "返回一个句柄到(from Microsoft Windows)一个对象的窗口。"
    hwnd = UserControl.hwnd
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "返回一个句柄(从 Microsoft Windows)到对象的设备上下文。"
    hDC = UserControl.hDC
End Property

Private Sub y16_Click()
    UserControl_Click
End Sub

Private Sub y8_Click()
    UserControl_Click
End Sub
