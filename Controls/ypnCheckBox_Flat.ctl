VERSION 5.00
Begin VB.UserControl ypnCheckBox_Flat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   FillStyle       =   0  'Solid
   ScaleHeight     =   975
   ScaleWidth      =   1410
   ToolboxBitmap   =   "ypnCheckBox_Flat.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   960
      Top             =   0
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   1
      Top             =   0
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   11
      Left            =   1200
      Picture         =   "ypnCheckBox_Flat.ctx":0312
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   10
      Left            =   960
      Picture         =   "ypnCheckBox_Flat.ctx":05FA
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   9
      Left            =   720
      Picture         =   "ypnCheckBox_Flat.ctx":08E2
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   8
      Left            =   480
      Picture         =   "ypnCheckBox_Flat.ctx":0BCA
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   7
      Left            =   240
      Picture         =   "ypnCheckBox_Flat.ctx":0EB2
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   6
      Left            =   0
      Picture         =   "ypnCheckBox_Flat.ctx":119A
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   5
      Left            =   1200
      Picture         =   "ypnCheckBox_Flat.ctx":1482
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   4
      Left            =   960
      Picture         =   "ypnCheckBox_Flat.ctx":176A
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   3
      Left            =   720
      Picture         =   "ypnCheckBox_Flat.ctx":1A52
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   2
      Left            =   480
      Picture         =   "ypnCheckBox_Flat.ctx":1D3A
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   1
      Left            =   240
      Picture         =   "ypnCheckBox_Flat.ctx":2022
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   0
      Left            =   0
      Picture         =   "ypnCheckBox_Flat.ctx":230A
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "ypnCheckBox_Flat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ypnCheckBox_Flat
' Author    : YPN
' Date      : 2018-03-24 00:09
' Purpose   : ±âÆ½»¯µÄcheckbox
'---------------------------------------------------------------------------------------

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long

Private Type POINT_API
    X As Long
    Y As Long
End Type

Public Enum State
    Unchecked = 0
    Checked = 1
    Mixed = 2
End Enum

Dim m_Font As Font
Dim m_Value As State
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR

Const m_def_Value = State.Unchecked
Const m_def_BackColor = vbButtonFace
Const m_def_ForeColor = vbBlack

Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub lbl_Click()
    Call UserControl_Click
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, lbl.Left, lbl.Top)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub p_Click()
    UserControl_Click
End Sub

Private Sub p_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub p_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
    Dim pnt As POINT_API
    UserControl.ScaleMode = 3
    GetCursorPos pnt
    ScreenToClient UserControl.hwnd, pnt
    
    If pnt.X < UserControl.ScaleLeft Or _
        pnt.Y < UserControl.ScaleTop Or _
        pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
        pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        
        define_pic
        Timer1.Enabled = False
        RaiseEvent MouseOut
    End If
End Sub

Private Sub UserControl_Click()
    If Value = Checked Then
        Value = Unchecked
    ElseIf Value = Unchecked Then
        Value = Checked
    ElseIf Value = Mixed Then
        Value = Unchecked
    End If
    define_pic
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    define_pic
    UserControl_Resize
    UserControl.BackColor = m_BackColor
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If Enabled = False Then
        enabled_pic
    Else: define_pic
    End If
    If Enabled = True Then lbl.ForeColor = m_ForeColor Else lbl.ForeColor = RGB(161, 161, 146)
End Property

Private Sub UserControl_InitProperties()
    Caption = Ambient.DisplayName
    Enabled = True
    Value = Unchecked
    Set Font = UserControl.Ambient.Font
    BackColor = m_def_BackColor
    ForeColor = m_def_ForeColor
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = True Then
        If Value = Checked Then
            p.Picture = Image1(6).Picture
        ElseIf Value = Mixed Then
            p.Picture = Image1(10).Picture
        ElseIf Value = Unchecked Then
            p.Picture = Image1(2).Picture
        End If
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X >= 0 And Y >= 0 And _
        X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
        If Button = vbLeftButton Then
            If Enabled = True Then
                If Value = Checked Then
                    p.Picture = Image1(6).Picture
                ElseIf Value = Mixed Then
                    p.Picture = Image1(10).Picture
                ElseIf Value = Unchecked Then
                    p.Picture = Image1(2).Picture
                End If
            End If
        Else
            If Enabled = True Then
                If Value = Checked Then
                    p.Picture = Image1(5).Picture
                ElseIf Value = Mixed Then
                    p.Picture = Image1(9).Picture
                ElseIf Value = Unchecked Then
                    p.Picture = Image1(1).Picture
                End If
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Value = PropBag.ReadProperty("Value", m_def_Value)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
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

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Caption", lbl.Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
End Sub

Public Property Get Value() As State
    Value = m_Value
End Property

Public Property Let Value(ByVal vNewValue As State)
    m_Value = vNewValue
    define_pic
    PropertyChanged "Value"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    lbl.Caption() = vNewCaption
    Call UserControl_Resize
    PropertyChanged "Caption"
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Set lbl.Font = m_Font
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Private Function define_pic()
    If Enabled = True Then
        If Value = Checked Then
            p.Picture = Image1(4).Picture
        ElseIf Value = Mixed Then
            p.Picture = Image1(8).Picture
        ElseIf Value = Unchecked Then
            p.Picture = Image1(0).Picture
        End If
    Else: enabled_pic
    End If
End Function

Private Function enabled_pic()
    If Value = Checked Then
        p.Picture = Image1(7).Picture
    ElseIf Value = Mixed Then
        p.Picture = Image1(11).Picture
    ElseIf Value = Unchecked Then
        p.Picture = Image1(3).Picture
    End If
End Function

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    If Enabled = True Then lbl.ForeColor = m_ForeColor Else lbl.ForeColor = RGB(161, 161, 146)
End Property
