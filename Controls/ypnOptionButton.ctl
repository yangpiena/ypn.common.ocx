VERSION 5.00
Begin VB.UserControl ypnOptionButton 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   FillStyle       =   0  'Solid
   ScaleHeight     =   840
   ScaleWidth      =   1920
   ToolboxBitmap   =   "ypnOptionButton.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   0
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   8
      Top             =   0
      Width           =   195
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   0
      Left            =   0
      Picture         =   "ypnOptionButton.ctx":0312
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   1
      Left            =   240
      Picture         =   "ypnOptionButton.ctx":05FA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   2
      Left            =   480
      Picture         =   "ypnOptionButton.ctx":08E2
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   3
      Left            =   720
      Picture         =   "ypnOptionButton.ctx":0BCA
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   4
      Left            =   960
      Picture         =   "ypnOptionButton.ctx":0EB2
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   5
      Left            =   1200
      Picture         =   "ypnOptionButton.ctx":119A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   6
      Left            =   1440
      Picture         =   "ypnOptionButton.ctx":1482
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   7
      Left            =   1680
      Picture         =   "ypnOptionButton.ctx":176A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Option1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "ypnOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ypnOptionButton
' Author    : YPN
' Date      : 2018-03-24 00:13
' Purpose   : ±âÆ½»¯µÄoptionButton
'---------------------------------------------------------------------------------------

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          'Aki
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long

Private Type POINT_API
    X As Long
    Y As Long
End Type

Dim mFont As Font
Dim mValue As Boolean
Dim mBackColor As OLE_COLOR
Dim mForeColor As OLE_COLOR
Dim mGroup As Boolean

Const defValue = False
Const defBackColor = vbButtonFace
Const defForeColor = vbBlack
Const defGroup = False

Dim mFrame As String
Dim chVal, btnDown As Integer
Dim mEnabled As Boolean 'I wrote this to make it more simple otherwise will jump all the time
'to sub Get Enabled.
'Also you can see that I used mValue instead of Value(the same reason)
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    btnDown = 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled = False Then Exit Sub
    If mValue = True Then p.Picture = img(6).Picture
    If mValue = False Then p.Picture = img(2).Picture
    btnDown = 1
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mEnabled = False Then Exit Sub
    If p.Picture = img(chVal).Picture Then Exit Sub
    If btnDown = 1 Then Exit Sub
    Timer1.Enabled = True
    
    If mEnabled = False Then
        If mValue = True Then p.Picture = img(6).Picture: chVal = 6
        If mValue = False Then p.Picture = img(2).Picture: chVal = 2
    Else
        If mValue = True Then p.Picture = img(5).Picture: chVal = 5
        If mValue = False Then p.Picture = img(1).Picture: chVal = 1
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub p_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, p.Left, p.Top)
End Sub

Private Sub p_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub p_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub p_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    RaiseEvent KeyPress(KeyAscii)
    RaiseEvent Click
    Call UserControl_MouseDown(1, 0, 0, 0)
End Sub

Private Sub p_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    RaiseEvent KeyUp(KeyCode, Shift)
    Call UserControl_Click
    btnDown = 0
End Sub

Private Sub lbl_Click()
    Call UserControl_Click
End Sub

Private Sub p_Click()
    UserControl_Click
End Sub
'********************************************************************************************
'I didn't comment in this project 'cause it is similar as XpCheckBox where are all
'the comments, ofcourse except this sub.
'Thanks to Mick Doherty for help in this sub.    Email:mdaudi100@ntlworld.com
'********************************************************************************************

Private Sub UserControl_Click()
    Dim OB As Object
    
    'Try to use variables and your program will run faster and it is much easier.
    'If you call Value it will jump to property all the time. But instead of that,
    'call mValue 'cause they have the same value and there is no jumping.
    
    If mValue = True Then CheckEnabled: Exit Sub 'Check Enabled must be here 'cause if user is
    'checking with keys picture will be changed.
    For Each OB In Parent.Controls
        If TypeOf OB Is ypnOptionButton Then
            If OB.Container Is Extender.Container Then
                If OB.Group = mGroup Then
                    OB.Value = False
                End If
            End If
        End If
    Next
    Value = True 'At the end give value to one that was clicked
    RaiseEvent Click
End Sub
Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
    UserControl_Resize
    DisablePc
    UserControl.BackColor = m_BackColor
    chVal = 1
    mGroup = False
End Sub

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
    Enabled = True
    mEnabled = True
    Value = False
    Set Font = UserControl.Ambient.Font
    BackColor = defBackColor
    ForeColor = defForeColor
End Sub

Private Sub UserControl_Resize()
    UserControl.ScaleMode = 1
    p.Height = 195
    p.Width = 195
    p.Left = 0
    p.Top = (UserControl.Height - p.Height) \ 2
    lbl.Top = (UserControl.Height - lbl.Height) \ 2
    lbl.Left = 240
End Sub

Private Function DisablePc()
    If mEnabled = True Then
        If mValue = True Then p.Picture = img(4).Picture
        If mValue = False Then p.Picture = img(0).Picture
    Else: EnablePc
    End If
End Function

Private Function EnablePc()
    If mValue = True Then p.Picture = img(7).Picture
    If mValue = False Then p.Picture = img(3).Picture
End Function

Private Sub CheckEnabled()
    If mEnabled = False Then EnablePc: lbl.ForeColor = &H80000011: Timer1.Enabled = False
    If mEnabled = True Then DisablePc: lbl.ForeColor = mForeColor
End Sub

Private Sub p_GotFocus()
    Call UserControl_MouseMove(0, 0, 0, 0)
    Timer1.Enabled = False
End Sub

Private Sub p_LostFocus()
    chVal = 7
    Call UserControl_MouseMove(0, 0, 0, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Value = PropBag.ReadProperty("Value", defValue)
    Group = PropBag.ReadProperty("Group", defGroup)
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    Caption = PropBag.ReadProperty("Caption", "Option1")
    BackColor = PropBag.ReadProperty("BackColor", defBackColor)
    ForeColor = PropBag.ReadProperty("ForeColor", defForeColor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Value", mValue, defValue)
    Call PropBag.WriteProperty("Group", mGroup, defGroup)
    Call PropBag.WriteProperty("Font", mFont, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "Option")
    Call PropBag.WriteProperty("BackColor", mBackColor, defBackColor)
    Call PropBag.WriteProperty("ForeColor", mForeColor, defForeColor)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
    mEnabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
    UserControl.Enabled() = NewEnabled
    mEnabled = NewEnabled
    CheckEnabled
    PropertyChanged "Enabled"
End Property
'**********************************************************************************************
Public Property Get Group() As Boolean
    Group = mGroup
End Property

Private Property Let Group(ByVal NewGroup As Boolean) 'User will not see this property.
    mGroup = NewGroup                                                 'That's the way it should be
End Property
'***********************************************************************************************
Public Property Get Value() As Boolean
    Value = mValue
End Property

Public Property Let Value(ByVal newValue As Boolean)
    mValue = newValue
    DisablePc
    PropertyChanged "Value"
End Property

Public Property Get Font() As Font
    Set Font = mFont
End Property

Public Property Set Font(ByVal newFont As Font)
    Set mFont = newFont
    Set UserControl.Font = newFont
    Set lbl.Font = mFont
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    lbl.Caption() = NewCaption
    Call UserControl_Resize
    PropertyChanged "Caption"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    mBackColor = NewBackColor
    UserControl.BackColor = mBackColor
    p.BackColor = mBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
    mForeColor = NewForeColor
    CheckEnabled
    PropertyChanged "ForeColor"
End Property

Private Sub Timer1_Timer()
    Dim dot As POINT_API
    UserControl.ScaleMode = 3
    Call GetCursorPos(dot)
    ScreenToClient UserControl.hwnd, dot
    
    If dot.X < UserControl.ScaleLeft Or _
        dot.Y < UserControl.ScaleTop Or _
        dot.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
        dot.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        
        If btnDown = 1 Then Exit Sub
        
        DisablePc
        Timer1.Enabled = False
        RaiseEvent MouseOut
    End If
End Sub
