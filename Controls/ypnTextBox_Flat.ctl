VERSION 5.00
Begin VB.UserControl ypnTextBox_Flat 
   Appearance      =   0  'Flat
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2610
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   1020
   ScaleWidth      =   2610
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D05C28&
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Text            =   "Super TextBox"
      Top             =   30
      Width           =   2295
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Caption         =   "Header"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Super TextBox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D05C28&
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D05C28&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "ypnTextBox_Flat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Properties variables
Dim ctlNumberBox As Boolean
Dim ctlVAlign As Byte
Dim ctlSelOnFocus As Boolean

' Events declaration
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event Click()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)

Enum VerticalAlignCostants  ' Defines the vertical alignement costants
    Top
    Center
    Bottom
End Enum

' *********************************************************************************
' EVENTS
' *********************************************************************************

Private Sub txt_Change()
    lbl.Caption = txt.Text
    RaiseEvent Change
End Sub

Private Sub txt_Click()
    RaiseEvent Click
End Sub

Private Sub txt_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txt_GotFocus()
    ' If the property SelOnFocus has been selected, then all text is selected
    If ctlSelOnFocus Then
        txt.SelStart = 0
        txt.SelLength = Len(txt)
    End If
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    ' If the property NumberBox has been selected, then only numbers can be typed
    If ctlNumberBox Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub

Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


' *********************************************************************************
' PROPERTIES
' *********************************************************************************

Public Property Get Font() As Font
Attribute Font.VB_Description = "The box font"
    Set Font = txt.Font
End Property

Public Property Set Font(ByRef newFont As Font)
    Set txt.Font = newFont
    Set lbl.Font = txt.Font
    SetAlign ctlVAlign
    PropertyChanged "FONT"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Text color"
    ForeColor = txt.ForeColor
End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
    txt.ForeColor = theCol
    lbl.ForeColor = theCol
    PropertyChanged "ForeColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Border color"
    BorderColor = shpBorder.BorderColor
End Property

Public Property Let BorderColor(ByVal theCol As OLE_COLOR)
    shpBorder.BorderColor = theCol
    PropertyChanged "BorderColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Text back color"
    BackColor = shpBorder.BackColor
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
    shpBorder.BackColor = theCol
    txt.BackColor = theCol
    PropertyChanged "BackColor"
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Guess what..."
    Text = txt.Text
End Property

Public Property Let Text(ByVal newText As String)
    txt.Text = newText
    lbl.Caption = newText
    PropertyChanged "Text"
End Property

Public Property Get AlignementHorizontal() As AlignmentConstants
Attribute AlignementHorizontal.VB_Description = "Text horizonal alignement"
    AlignementHorizontal = txt.Alignment
End Property

Public Property Let AlignementHorizontal(ByVal newValue As AlignmentConstants)
    If newValue > 2 Then newValue = 2
    txt.Alignment = newValue
    lbl.Alignment = txt.Alignment
    PropertyChanged ("AlignementHorizontal")
End Property

Public Property Get AlignementVertical() As VerticalAlignCostants
Attribute AlignementVertical.VB_Description = "Text vertical alignement"
    AlignementVertical = ctlVAlign
End Property

Public Property Let AlignementVertical(ByVal newValue As VerticalAlignCostants)
    If newValue > 2 Then newValue = 2
    ctlVAlign = newValue
    SetAlign ctlVAlign
    PropertyChanged ("AlignementVertical")
End Property

Public Property Get NumberBox() As Boolean
Attribute NumberBox.VB_Description = "If TRUE, only numbers are allowed in the box"
    NumberBox = ctlNumberBox
End Property

Public Property Let NumberBox(ByVal newValue As Boolean)
    ctlNumberBox = newValue
    PropertyChanged ("NumberBox")
End Property

Public Property Get SelOnFocus() As Boolean
Attribute SelOnFocus.VB_Description = "If TRUE, when the box gets focus all text inside is selected"
    SelOnFocus = ctlSelOnFocus
End Property

Public Property Let SelOnFocus(ByVal newValue As Boolean)
    ctlSelOnFocus = newValue
    PropertyChanged ("SelOnFocus")
End Property

Public Property Get LabelBox() As Boolean
Attribute LabelBox.VB_Description = "If TRUE, this control becomes a normal label"
    LabelBox = lbl.Visible
End Property

Public Property Let LabelBox(ByVal newValue As Boolean)
    lbl.Visible = newValue
    txt.Visible = Not newValue
    PropertyChanged ("LabelBox")
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Enable/disable the box"
    Enabled = txt.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    txt.Enabled = newValue
    lbl.Enabled = newValue
    PropertyChanged ("Enabled")
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Enable/disable typing in the text box"
    Locked = txt.Locked
End Property

Public Property Let Locked(ByVal newValue As Boolean)
    txt.Locked = newValue
    PropertyChanged ("Locked")
End Property

Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Character to show to hide typed ones"
    PasswordChar = txt.PasswordChar
End Property

Public Property Let PasswordChar(ByVal newValue As String)
    If Len(newValue) > 1 Then newValue = Left(newValue, 1)
    txt.PasswordChar = newValue
    PropertyChanged ("PasswordChar")
End Property

Public Property Get Header() As Boolean
Attribute Header.VB_Description = "Show/hide the header"
    Header = lblHeader.Visible
End Property

Public Property Let Header(ByVal newValue As Boolean)
    lblHeader.Visible = newValue
    SetAlign ctlVAlign
    PropertyChanged ("Header")
End Property

Public Property Get HeaderAlignement() As AlignmentConstants
Attribute HeaderAlignement.VB_Description = "Header alignement"
    HeaderAlignement = lblHeader.Alignment
End Property

Public Property Let HeaderAlignement(ByVal newValue As AlignmentConstants)
    If newValue > 2 Then newValue = 2
    lblHeader.Alignment = newValue
    PropertyChanged ("HeaderAlignement")
End Property

Public Property Get HeaderForeColor() As OLE_COLOR
Attribute HeaderForeColor.VB_Description = "Header color"
    HeaderForeColor = lblHeader.ForeColor
End Property

Public Property Let HeaderForeColor(ByVal newValue As OLE_COLOR)
    lblHeader.ForeColor = newValue
    PropertyChanged ("HeaderForeColor")
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
Attribute HeaderBackColor.VB_Description = "Header back color"
    HeaderBackColor = lblHeader.BackColor
End Property

Public Property Let HeaderBackColor(ByVal newValue As OLE_COLOR)
    lblHeader.BackColor = newValue
    UserControl.BackColor = newValue
    PropertyChanged ("HeaderBackColor")
End Property

Public Property Get HeaderFont() As Font
Attribute HeaderFont.VB_Description = "Header font"
    Set HeaderFont = lblHeader.Font
End Property

Public Property Set HeaderFont(ByRef newFont As Font)
    Set lblHeader.Font = newFont
    SetAlign ctlVAlign
    PropertyChanged "HeaderFont"
End Property

Public Property Get HeaderCaption() As String
Attribute HeaderCaption.VB_Description = "Header caption (if used)"
    HeaderCaption = lblHeader.Caption
End Property

Public Property Let HeaderCaption(ByVal newValue As String)
    lblHeader.Caption = newValue
    PropertyChanged ("HeaderCaption")
End Property


' *********************************************************************************
' USER CONTROL
' *********************************************************************************

Private Sub UserControl_InitProperties()
    ctlNumberBox = False
    ctlSelOnFocus = True
    ctlVAlign = 1
End Sub

' Resize text, label, header and border to the control size
Private Sub UserControl_Resize()
    If UserControl.Width < 300 Then UserControl.Width = 300
    If UserControl.Height < 60 Then UserControl.Width = 60
    shpBorder.Width = UserControl.Width
    txt.Width = shpBorder.Width - 90
    SetAlign ctlVAlign
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txt.Text = PropBag.ReadProperty("Text", "")
    lbl.Caption = txt.Text
    txt.Alignment = PropBag.ReadProperty("AlignementHorizontal", 1)
    lbl.Alignment = txt.Alignment
    ctlNumberBox = PropBag.ReadProperty("NumberBox", 0)
    ctlVAlign = PropBag.ReadProperty("AlignementVertical", 1)
    SetAlign ctlVAlign
    txt.Locked = PropBag.ReadProperty("Locked", 0)
    txt.Enabled = PropBag.ReadProperty("Enabled", 1)
    lbl.Enabled = txt.Enabled
    txt.ForeColor = PropBag.ReadProperty("ForeColor", &HD05C28)
    lbl.ForeColor = txt.ForeColor
    shpBorder.BorderColor = PropBag.ReadProperty("BorderColor", &HD05C28)
    shpBorder.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    txt.BackColor = shpBorder.BackColor
    lbl.BackColor = shpBorder.BackColor
    Set txt.Font = PropBag.ReadProperty("FONT", "Arial")
    Set lbl.Font = txt.Font
    txt.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    lbl.Visible = PropBag.ReadProperty("LabelBox", 0)
    txt.Visible = Not lbl.Visible
    ctlSelOnFocus = PropBag.ReadProperty("SelOnFocus", 1)
    lblHeader.Visible = PropBag.ReadProperty("Header", 0)
    lblHeader.Alignment = PropBag.ReadProperty("HeaderAlignement", 1)
    lblHeader.ForeColor = PropBag.ReadProperty("HeaderForeColor", &HD05C28)
    lblHeader.BackColor = PropBag.ReadProperty("HeaderBackColor", &HFFFFFF)
    UserControl.BackColor = lblHeader.BackColor
    Set lblHeader.Font = PropBag.ReadProperty("HeaderFont", "Verdana")
    lblHeader.Caption = PropBag.ReadProperty("HeaderCaption", "Header")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Text", txt.Text
    PropBag.WriteProperty "AlignementHorizontal", txt.Alignment
    PropBag.WriteProperty "NumberBox", ctlNumberBox
    PropBag.WriteProperty "AlignementVertical", ctlVAlign
    PropBag.WriteProperty "Locked", txt.Locked
    PropBag.WriteProperty "Enabled", txt.Enabled
    PropBag.WriteProperty "ForeColor", txt.ForeColor
    PropBag.WriteProperty "BorderColor", shpBorder.BorderColor
    PropBag.WriteProperty "BackColor", shpBorder.BackColor
    PropBag.WriteProperty "FONT", txt.Font
    PropBag.WriteProperty "LabelBox", lbl.Visible
    PropBag.WriteProperty "SelOnFocus", ctlSelOnFocus
    PropBag.WriteProperty "Header", lblHeader.Visible
    PropBag.WriteProperty "HeaderAlignement", lblHeader.Alignment
    PropBag.WriteProperty "HeaderForeColor", lblHeader.ForeColor
    PropBag.WriteProperty "HeaderBackColor", lblHeader.BackColor
    PropBag.WriteProperty "HeaderFont", lblHeader.Font
    PropBag.WriteProperty "HeaderCaption", lblHeader.Caption
End Sub


' *********************************************************************************
' CUSTOM ROUTINES
' *********************************************************************************

' This routine is used to align the controls
Private Sub SetAlign(ByVal Value As Byte)
Dim yOffset As Integer
    lblHeader.AutoSize = True   ' Automatically gets the header minimum height
    lblHeader.AutoSize = False
    If lblHeader.Visible Then   ' If the header is visible, all other things get shifted down
        yOffset = lblHeader.Height + 15
    Else
        yOffset = 0
    End If
    lblHeader.Left = 0          ' Ensure that the header is positioned in the left side
    lblHeader.Top = 0           ' Ensure that the header is positioned in the top side
    lblHeader.Width = shpBorder.Width
    shpBorder.Top = yOffset
    shpBorder.Height = UserControl.Height - IIf(lblHeader.Visible, yOffset, 0)
    lbl.AutoSize = True         ' It automatically gets the label minimum height
    lbl.AutoSize = False
    Select Case Value
        Case 0                  ' TOP
            txt.Top = 15 + shpBorder.Top
            lbl.Top = 15 + shpBorder.Top
        Case 1                  ' CENTER
            lbl.Top = shpBorder.Top + ((shpBorder.Height - lbl.Height) / 2)
            txt.Top = lbl.Top
        Case 2                  ' BOTTOM
            lbl.Top = shpBorder.Top + (shpBorder.Height - lbl.Height) - 15
            txt.Top = lbl.Top
    End Select
    lbl.Width = txt.Width       ' Readjust the width
    lbl.Left = txt.Left         ' Readjust the position
    txt.Height = 0              ' Automatically gets the text minimum height
End Sub

Sub About()
Attribute About.VB_Description = "Informazioni sul MIO codice"
Attribute About.VB_UserMemId = -552
    On Error Resume Next
    MsgBox "Super TextBox" & vbCrLf & vbCrLf & "Written by:" & vbCrLf & "Umberto Nocentini" & vbCrLf & vbCrLf & "You can freely use this control as you want!", vbInformation
End Sub
