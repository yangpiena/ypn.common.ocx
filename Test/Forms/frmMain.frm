VERSION 5.00
Object = "{B84B182A-14BC-41FC-856A-D22D2973E52D}#3.1#0"; "ypn.common.ocx.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8850
   StartUpPosition =   3  '窗口缺省
   Begin YPNCommonOCX.ypnButton_Shape ypnButton_Shape2 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "测试Tab"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin YPNCommonOCX.ypnButton_Shape ypnButton_Shape1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "测试扁平化"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ypnButton_Shape1_Click()
    frmFlat.Show
End Sub

Private Sub ypnButton_Shape2_Click()
    frmTab.Show
End Sub

