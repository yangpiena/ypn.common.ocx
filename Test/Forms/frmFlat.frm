VERSION 5.00
Object = "{B84B182A-14BC-41FC-856A-D22D2973E52D}#3.2#0"; "ypn.common.ocx.ocx"
Begin VB.Form frmFlat 
   BorderStyle     =   0  'None
   Caption         =   "Flat"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   1800
   End
   Begin YPNCommonOCX.ypnImage_Flat ypnImage_Flat2 
      Height          =   960
      Left            =   360
      TabIndex        =   17
      Top             =   960
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Picture         =   "frmFlat.frx":0000
      PictureHover    =   "frmFlat.frx":027E
      PictureDown     =   "frmFlat.frx":05C6
   End
   Begin YPNCommonOCX.ypnShape_Flat ypnShape_Flat3 
      Height          =   495
      Left            =   4080
      TabIndex        =   16
      Top             =   6120
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Picture_Normal  =   "frmFlat.frx":07E8
      Picture_Down    =   "frmFlat.frx":0804
      Picture_Hover   =   "frmFlat.frx":0820
      Stretch         =   0   'False
      Caption         =   "确 定"
      BackGround      =   15725042
      BackColorNormal =   15725042
      BackColorHover  =   14737632
      BackColorDown   =   12632064
      BorderColorNormal=   11776947
      BorderColorHover=   11776947
      BorderColorDown =   12632064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      Style           =   0
      ForeColorNormal =   0
      ForeColorHover  =   0
      ForeColorDown   =   0
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin YPNCommonOCX.ypnShape_Flat ypnShape_Flat2 
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   6120
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Picture_Normal  =   "frmFlat.frx":083C
      Picture_Down    =   "frmFlat.frx":0858
      Picture_Hover   =   "frmFlat.frx":0874
      Stretch         =   0   'False
      Caption         =   "按 钮"
      BackGround      =   15725042
      BackColorNormal =   15725042
      BackColorHover  =   14737632
      BackColorDown   =   16777215
      BorderColorNormal=   12632064
      BorderColorHover=   11776947
      BorderColorDown =   11776947
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      ForeColorNormal =   0
      ForeColorHover  =   0
      ForeColorDown   =   0
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin YPNCommonOCX.ypnShape_Flat ypnShape_Flat1 
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   6120
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      Picture_Normal  =   "frmFlat.frx":0890
      Picture_Down    =   "frmFlat.frx":08AC
      Picture_Hover   =   "frmFlat.frx":08C8
      Stretch         =   0   'False
      Caption         =   "开 始"
      BackGround      =   15725042
      BackColorNormal =   15725042
      BackColorHover  =   14737632
      BackColorDown   =   16777215
      BorderColorNormal=   32768
      BorderColorHover=   11776947
      BorderColorDown =   11776947
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderCustom    =   11776947
      ForeColorNormal =   0
      ForeColorHover  =   0
      ForeColorDown   =   0
      Text_Visible    =   -1  'True
      StretchToText   =   0   'False
   End
   Begin YPNCommonOCX.ypnListBox_Flat ypnListBox_Flat1 
      Height          =   1935
      Left            =   240
      TabIndex        =   13
      Top             =   3840
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3413
      Picture         =   "frmFlat.frx":08E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Path            =   "C:\WINDOWS\system32\"
   End
   Begin YPNCommonOCX.ypnCheckBox_Flat ypnCheckBox_Flat3 
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   2
   End
   Begin YPNCommonOCX.ypnCheckBox_Flat ypnCheckBox_Flat2 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   1
   End
   Begin YPNCommonOCX.ypnCheckBox_Flat ypnCheckBox_Flat1 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin YPNCommonOCX.ypnOptionButton ypnOptionButton3 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Enabled         =   0   'False
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ypnOptionButton3"
   End
   Begin YPNCommonOCX.ypnOptionButton ypnOptionButton2 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ypnOptionButton2"
   End
   Begin YPNCommonOCX.ypnOptionButton ypnOptionButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ypnOptionButton1"
   End
   Begin YPNCommonOCX.ypnComboBox_Flat ypnComboBox_Flat2 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxListLength   =   -1
      NormalBorderColor=   16576
      NormalColorText =   16576
      NumberItemsToShow=   -1
      ShadowColorText =   6582129
   End
   Begin YPNCommonOCX.ypnTextBox_Flat ypnTextBox_Flat2 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Text            =   "Super TextBox"
      AlignementHorizontal=   2
      NumberBox       =   0   'False
      AlignementVertical=   1
      Locked          =   0   'False
      Enabled         =   -1  'True
      ForeColor       =   13655080
      BorderColor     =   13655080
      BackColor       =   16777215
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelBox        =   0   'False
      SelOnFocus      =   -1  'True
      Header          =   0   'False
      HeaderAlignement=   2
      HeaderForeColor =   -2147483640
      HeaderBackColor =   -2147483633
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderCaption   =   "Header"
   End
   Begin YPNCommonOCX.ypnTextBox_Flat ypnTextBox_Flat1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Text            =   "Super TextBox"
      AlignementHorizontal=   2
      NumberBox       =   0   'False
      AlignementVertical=   1
      Locked          =   0   'False
      Enabled         =   -1  'True
      ForeColor       =   13655080
      BorderColor     =   13655080
      BackColor       =   16777215
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelBox        =   0   'False
      SelOnFocus      =   -1  'True
      Header          =   0   'False
      HeaderAlignement=   2
      HeaderForeColor =   -2147483640
      HeaderBackColor =   -2147483633
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderCaption   =   "Header"
   End
   Begin YPNCommonOCX.ypnComboBox_Flat ypnComboBox_Flat1 
      Height          =   420
      Left            =   1440
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   741
      BackColor       =   13430215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxListLength   =   -1
      NumberItemsToShow=   -1
      OfficeAppearance=   2
      ShadowColorText =   6582129
   End
   Begin YPNCommonOCX.ypnProgressBar_Flat ypnProgressBar_Flat1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
   End
   Begin YPNCommonOCX.ypnImage_Flat ypnImage_Flat1 
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Top             =   0
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   503
      Picture         =   "frmFlat.frx":0900
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "测试扁平化"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   10
      Height          =   615
      Left            =   120
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmFlat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As String

Private Sub Form_Load()
    i = 1
    Shape1.Height = Me.Height - 2
    Shape1.Width = Me.Width - 2
    Shape1.Top = 1
    Shape1.Left = 1
    ypnProgressBar_Flat1.Value = 50
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDrag Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoDrag Me
End Sub

Private Sub Timer1_Timer()
    i = i + 1
    ypnProgressBar_Flat1.Value = i
End Sub

Private Sub ypnImage_Flat1_Click()
    Unload Me
End Sub

Private Sub ypnShape_Flat3_Click()
    ypnProgressBar_Flat1.Value = 1
    i = 1
End Sub
