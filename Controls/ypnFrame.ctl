VERSION 5.00
Begin VB.UserControl ypnFrame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4125
   ScaleWidth      =   4170
   ToolboxBitmap   =   "ypnFrame.ctx":0000
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   480
      ScaleHeight     =   405
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Shape Border 
      Height          =   405
      Left            =   870
      Top             =   0
      Width           =   345
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   1380
      X2              =   1380
      Y1              =   0
      Y2              =   270
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   1290
      X2              =   1290
      Y1              =   0
      Y2              =   270
   End
   Begin VB.Image imgPicture 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   30
      Stretch         =   -1  'True
      Top             =   30
      Width           =   195
   End
End
Attribute VB_Name = "ypnFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const m_def_ColorTop = vbWhite
Const m_def_ColorBottom = &HF3E9DA
Const m_def_Caption = "Jay's Frame"
Const m_def_Enabled = 0

Dim PictureExists As Boolean
Dim m_TransParent As Boolean
Dim m_borderVisible As Boolean

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long _
    ) As Long

Private Const SRCCOPY = &HCC0020
Private Const SRCAND = &H8800C6
Private Const SRCPAINT = &HEE0086
Private Const SRCINVERT = &H660046

'Dim m_Picture As Picture
Dim m_ColorTop As OLE_COLOR
Dim m_ColorBottom As OLE_COLOR
Dim m_Caption As String
Dim m_Enabled As Boolean

Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseExit()
Event MouseIn()

Private Sub UserControl_Resize()
  RedrawPanel
End Sub

Private Sub RedrawPanel()

     Dim Temp As tpAPI_RECT

     If m_borderVisible Then
          Border.Visible = True
          Line1.Visible = True
          
     Else
          Border.Visible = False
          Line1.Visible = False
Line2.Visible = False
     End If

     If m_TransParent Then
          UserControl.Cls
          CopyBackGround UserControl.Parent.hWnd, UserControl.hWnd, UserControl.hDC
          Line1.X1 = 0
          Line1.X2 = ScaleWidth
          Line1.Y1 = PixelsToTwips_height(Temp.lBottom)
          Line1.Y2 = PixelsToTwips_height(Temp.lBottom)
          Border.Move 0, 0, ScaleWidth, ScaleHeight
          UserControl.Refresh
     Else
  
          Temp.lTop = 0
          Temp.lLeft = 0
          Temp.lRight = TwipsToPixels_width(ScaleWidth)
          Temp.lBottom = TwipsToPixels_height(270)
          Call DrawTopDownGradient(UserControl.hDC, Temp, m_ColorTop, m_ColorBottom)
  
          Line1.X1 = 0
          Line1.X2 = ScaleWidth
          Line1.Y1 = PixelsToTwips_height(Temp.lBottom)
          Line1.Y2 = PixelsToTwips_height(Temp.lBottom)
  
          Temp.lTop = TwipsToPixels_height(270) - 1
          Temp.lLeft = 0
          Temp.lRight = TwipsToPixels_width(ScaleWidth)
          Temp.lBottom = TwipsToPixels_height(ScaleHeight)
          Call DrawTopDownGradient(UserControl.hDC, Temp, m_ColorTop, m_ColorBottom)
  
          Border.Move 0, 0, ScaleWidth, ScaleHeight
  
          UserControl.FontBold = True
   
          UserControl.CurrentY = (Line1.Y1 / 2) - (UserControl.TextHeight(")") / 2) + 10

          If PictureExists Then
               UserControl.CurrentX = 130 + imgPicture.Width + imgPicture.Left
          Else
               UserControl.CurrentX = 100

          End If

     End If

     UserControl.Print m_Caption

End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
  ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  UserControl.ForeColor() = New_ForeColor
  RedrawPanel
  PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  PropertyChanged "Font"
End Property



Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
    If (X < 0) Or (Y < 0) Or (X > UserControl.Width) Or (Y > UserControl.Height) Then
      ReleaseCapture
      RaiseEvent MouseExit  ' 鼠标离开的代码
    Else
      SetCapture hWnd
      RaiseEvent MouseIn    ' 鼠标进入的代码
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
  BorderColor = Border.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
  Border.BorderColor() = New_BorderColor
  Line1.BorderColor() = New_BorderColor
  PropertyChanged "BorderColor"
End Property

Private Sub UserControl_InitProperties()
  m_Enabled = m_def_Enabled
  m_Caption = m_def_Caption
  m_ColorTop = m_def_ColorTop
  m_ColorBottom = m_def_ColorBottom
'  Set m_Picture = LoadPicture("")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
  m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
  Border.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
  Line1.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
  m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
  m_ColorTop = PropBag.ReadProperty("ColorTop", m_def_ColorTop)
  m_ColorBottom = PropBag.ReadProperty("ColorBottom", m_def_ColorBottom)
  m_borderVisible = PropBag.ReadProperty("borderVisible", True)
  m_TransParent = PropBag.ReadProperty("Transparent", False)
'  Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
  Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

     Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
     Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
     Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
     Call PropBag.WriteProperty("BorderColor", Border.BorderColor, -2147483640)
     Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
     Call PropBag.WriteProperty("ColorTop", m_ColorTop, m_def_ColorTop)
     Call PropBag.WriteProperty("ColorBottom", m_ColorBottom, m_def_ColorBottom)
     Call PropBag.WriteProperty("borderVisible", m_borderVisible, True)
     Call PropBag.WriteProperty("Transparent", m_TransParent, False)
     '  Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
     Call PropBag.WriteProperty("Picture", Picture, Nothing)

End Sub

Public Property Get Caption() As String
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  m_Caption = New_Caption
  PropertyChanged "Caption"
  RedrawPanel
End Property

Public Property Get ColorTop() As OLE_COLOR
  ColorTop = m_ColorTop
End Property

Public Property Let ColorTop(ByVal New_ColorTop As OLE_COLOR)
  m_ColorTop = New_ColorTop
  RedrawPanel
  PropertyChanged "ColorTop"
End Property

Public Property Get ColorBottom() As OLE_COLOR
  ColorBottom = m_ColorBottom
End Property

Public Property Let ColorBottom(ByVal New_ColorBottom As OLE_COLOR)
  m_ColorBottom = New_ColorBottom
  RedrawPanel
  PropertyChanged "ColorBottom"
End Property

Private Function PixelsToTwips_height(pxls)
    PixelsToTwips_height = pxls * Screen.TwipsPerPixelY
End Function

Private Function PixelsToTwips_width(pxls)
    PixelsToTwips_width = pxls * Screen.TwipsPerPixelX
End Function

Private Function TwipsToPixels_height(pxls)
    TwipsToPixels_height = pxls \ Screen.TwipsPerPixelY
End Function

Private Function TwipsToPixels_width(pxls)
    TwipsToPixels_width = pxls \ Screen.TwipsPerPixelX
End Function

'Public Property Get Picture() As Picture
'  Set Picture = m_Picture
'End Property

'Public Property Set Picture(ByVal New_Picture As Picture)
'  Set m_Picture = New_Picture
'  PropertyChanged "Picture"
'End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
  Set Picture = picPicture.Picture
End Property

Private Sub ResizeToIcon()
  Dim Diff As Single
  Dim Y As Integer
  Y = 220
  If picPicture.Height > picPicture.Width Then
    Diff = Y / picPicture.Height
    imgPicture.Height = Y
    imgPicture.Width = picPicture.Width * Diff
  Else
    Diff = Y / picPicture.Width
    imgPicture.Width = Y
    imgPicture.Height = picPicture.Height * Diff
  End If
  
  
  imgPicture.Picture = picPicture.Image
  Line2.Y1 = 0
  Line2.Y2 = 270
  Line2.X1 = imgPicture.Width + imgPicture.Left + 30
  Line2.X2 = imgPicture.Width + imgPicture.Left + 30
  Line2.BorderColor = Border.BorderColor
End Sub

Public Property Set Picture(ByVal New_Picture As Picture)
  Set picPicture.Picture = New_Picture
  On Error GoTo Err
  If Val(New_Picture) > 0 Then
    PictureExists = True
    ResizeToIcon
  Else
    PictureExists = False
  End If
  imgPicture.Visible = PictureExists
  Line2.Visible = PictureExists
  PropertyChanged "Picture"
  GoTo Err2
  
Err:
  PictureExists = False
  imgPicture.Visible = PictureExists
  Line2.Visible = PictureExists
  PropertyChanged "Picture"
Err2:
  RedrawPanel
End Property

Public Property Get TransParent() As Boolean
      TransParent = m_TransParent
End Property

Public Property Let TransParent(ByVal vNewValue As Boolean)
       m_TransParent = vNewValue
       RedrawPanel
End Property

Public Property Get borderVisible() As Boolean
borderVisible = m_borderVisible
End Property

Public Property Let borderVisible(ByVal vNewValue As Boolean)
m_borderVisible = vNewValue
PropertyChanged "borderVisible"
RedrawPanel
End Property
